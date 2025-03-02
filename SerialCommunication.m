classdef SerialCommunication
    properties (Access = public)
        SerialCOM  % Serial port object
    end
    %% 
    methods (Access = public)
    %% 
        % Constructor
        function obj = SerialCommunication(port, baudRate, dataBits, stopBits, parity, handshake)
            % Check if the specific port is open and close it
            openPorts = instrfind('Port', port);
            if ~isempty(openPorts)
                fclose(openPorts);
                delete(openPorts);
            end
            
            % Create serial port object with settings
            obj.SerialCOM = serial(port);
            set(obj.SerialCOM, 'BaudRate', baudRate);
            set(obj.SerialCOM, 'DataBits', dataBits);
            set(obj.SerialCOM, 'Parity', parity);
            set(obj.SerialCOM, 'StopBits', stopBits);
            set(obj.SerialCOM, 'ReadAsyncMode', 'continuous');
            set(obj.SerialCOM, 'InputBufferSize', 35000); % Default buffer size
            
            % Open the serial port
            fopen(obj.SerialCOM);
        end
        %% 
        
        % Close port
        function closePort(obj)
            if ~isempty(obj.SerialCOM) && isvalid(obj.SerialCOM)
                fclose(obj.SerialCOM);
                delete(obj.SerialCOM);
            end
        end
        %% 
        
        % Take measurement
        function takeMeasurement(obj, axesHandle, clearPlot, PlotCount)
            try
                disp(['takeMeasurement called with ' num2str(nargin) ' arguments']);
                disp(['Arguments received: axesHandle type=' class(axesHandle) ...
                      ', clearPlot=' num2str(clearPlot)]);

                % Start single measurement
                fprintf(obj.SerialCOM, 'Start single');
                wait = waitbar(1/3, 'Measuring...');
                
                % Wait for measurement to complete based on exposure time
                pause(6); % Default wait time for auto exposure
                
                % Request spectrum data
                waitbar(2/3, wait, 'Get current spectrum...');
                fprintf(obj.SerialCOM, 'Get spectrum cur');
                
                % Wait for data to be ready
                pause(3);
                
                % Parse the spectrum data
                [wavelength, value] = obj.parseSpectrumData();
                
                % Plot the data with legend
                if ~isempty(wavelength) && ~isempty(value)
                    obj.plotSpectrum(wavelength, value, axesHandle, clearPlot, PlotCount);
                end
                
                close(wait);
                
            catch ME
                if exist('wait', 'var')
                    close(wait);
                end
                disp(['Error occurred in function: ' ME.stack(1).name]);
                disp(['Line number: ' num2str(ME.stack(1).line)]);
                disp(['Error message: ' ME.message]);
                error('Measurement failed: %s', ME.message);
            end
        end
        %% 
        
        % Parse spectrum data
        function [Wavelength, Value] = parseSpectrumData(obj)
            Wavelength = [];
            Value = [];
            w = 0;
            StartSpec = 0;
            EndSpec = 0;
            
            % Get initial number of bytes available
            N = obj.SerialCOM.BytesAvailable;
            
            while N > 0
                line = fgetl(obj.SerialCOM);
                
                if StartSpec == 0
                    if contains(line, '--- Current spectrum ----------------------------- start ---')
                        % Skip header lines
                        for i = 1:3
                            line = fgetl(obj.SerialCOM);
                        end
                        StartSpec = 1;
                    end
                elseif EndSpec ~= 1
                    if contains(line, '--- Current spectrum ------------------------------- end ---')
                        EndSpec = 1;
                    else
                        % Parse data lines
                        [wl, val] = strread(line, '%f%f', 'delimiter', ' ');
                        if ~isempty(wl) && ~isempty(val)
                            w = w + 1;
                            Wavelength(w) = wl;
                            Value(w) = val;
                        end
                    end
                end
                
                N = obj.SerialCOM.BytesAvailable;
            end
            
            if isempty(Wavelength) || isempty(Value)
                error('No valid data received from spectrometer');
            end
        end
        %% 

        % Plot spectrum
    function plotSpectrum(obj, wavelength, value, axesHandle, clearPlot, PlotCount)
    if nargin < 6
        PlotCount = 1;  % Default value if not provided
    end

    if isempty(wavelength) || isempty(value) || length(wavelength) ~= length(value)
        error('Invalid wavelength or value data for plotting.');
    end

    if clearPlot
        cla(axesHandle);
    end

    % Find peak wavelength (dominant color)
    [~, maxIndex] = max(value);
    peakWavelength = wavelength(maxIndex);
    
    % Get RGB color & color name
    colorName = obj.getColorDescription(wavelength, value);
     colorRGB = obj.getColorRGB(colorName);

    % Plot the spectrum with dynamic name
    semilogy(axesHandle, wavelength, value, 'Color', colorRGB, ...
        'LineWidth', 2, 'DisplayName', ['Line ' num2str(PlotCount) ' (' colorName ')']);
    
    grid(axesHandle, 'on');
    xlabel(axesHandle, 'Wavelength (nm)');
    ylabel(axesHandle, 'Intensity');
    
    % Enable legend with dynamic names
    legend(axesHandle, 'show', 'Location', 'best');
    
    xlim(axesHandle, [min(wavelength) max(wavelength)]);
    title(axesHandle, 'Spectrometer Data');
    end
%% 


function colorName = getColorDescription(~, wavelength, intensity)
    try
        % First smooth the intensity data to reduce noise
        windowSize = 15;
        smoothedIntensity = movmean(intensity, windowSize);
        
        % Define spectral ranges
        validRange = wavelength >= 250 & wavelength <= 750;
        
        % Work with smoothed data in valid range
        validIntensities = smoothedIntensity(validRange);
        validWavelengths = wavelength(validRange);
        
        % Find the maximum intensity and its wavelength
        [maxIntensity, maxIndex] = max(validIntensities);
        peakWavelength = validWavelengths(maxIndex);
        
        % Check for white LED characteristics
        bluePeakRange = validWavelengths >= 440 & validWavelengths <= 470;  % Blue peak range
        yellowPeakRange = validWavelengths >= 520 & validWavelengths <= 560;  % Yellow-green peak range
        
        % Find the intensity values in these ranges
        bluePeakIntensity = max(validIntensities(bluePeakRange));
        yellowPeakIntensity = max(validIntensities(yellowPeakRange));
        
        % Define a threshold for recognizing a white LED
        whiteLEDThreshold = 0.3;  % Relative intensity threshold for secondary peaks
        
        % Check if both peaks are detected and meet the threshold for white LED
        if bluePeakIntensity > maxIntensity * whiteLEDThreshold && yellowPeakIntensity > maxIntensity * whiteLEDThreshold
            colorName = 'White LED';
            disp('Detected White LED spectrum with characteristic blue and yellow peaks.');
        else
            % Proceed with color classification for other LEDs
            % Calculate weighted center of mass around the peak
            threshold = maxIntensity * 0.3; % Consider points above 30% of max
            significantPeaks = validIntensities > threshold;
            weightedWavelength = sum(validWavelengths(significantPeaks) .* validIntensities(significantPeaks)) / ...
                                sum(validIntensities(significantPeaks));
            
            % Print diagnostic information
            fprintf('Peak wavelength: %.2f nm\n', peakWavelength);
            fprintf('Weighted center wavelength: %.2f nm\n', weightedWavelength);
            
            % Color determination based on weighted center
            if weightedWavelength < 410
                colorName = 'UV/Purple';
            elseif weightedWavelength < 495
                colorName = 'Blue';
            elseif weightedWavelength < 570
                colorName = 'Green';
            elseif weightedWavelength < 590
                colorName = 'Yellow';
            elseif weightedWavelength < 620
                colorName = 'Orange';
            elseif weightedWavelength <= 750
                colorName = 'Red';
            else
                colorName = 'Unknown';
            end
            fprintf('Determined color: %s\n', colorName);
        end
        
    catch ME
        warning('Color detection failed: %s', ME.message);
        colorName = 'Unknown';
    end
end
        %%
        function colorRGB = getColorRGB(~, colorName)
    % Map color names to RGB values
    switch lower(colorName)
        case 'uv/purple'
            colorRGB = [0.5, 0, 1];  % Purple for UV
        case 'blue'
            colorRGB = [0, 0, 1];    % Pure Blue
        case 'green'
            colorRGB = [0, 1, 0];    % Pure Green
        case 'yellow'
            colorRGB = [1, 1, 0];    % Pure Yellow
        case 'orange'
            colorRGB = [1, 0.5, 0];  % Pure Orange
        case 'red'
            colorRGB = [1, 0, 0];    % Pure Red
        case 'white led'
            colorRGB = [0, 0, 0];    % Black for White LED (as in Spectrum.m)
        otherwise
            colorRGB = [0.5, 0.5, 0.5];  % Gray for unknown
    end
end
        %% 

        % Set auto exposure
        function setAutoExposure(obj, state)
            if isempty(obj.SerialCOM) || ~isvalid(obj.SerialCOM)
                error('Serial port is not open.');
            end
            if state
                fprintf(obj.SerialCOM, 'Set autoexposure on');
            else
                fprintf(obj.SerialCOM, 'Set autoexposure off');
            end
            pause(0.05);
        end
        %% 
        % Set exposure time
        function setExposureTime(obj, exposureTime)
            if isempty(obj.SerialCOM) || ~isvalid(obj.SerialCOM)
                error('Serial port is not open.');
            end
            fprintf(obj.SerialCOM, ['Set exposuretime ' num2str(exposureTime)]);
            pause(0.05);
        end
        %% 
        % Save spectrum data to CSV file
        function saveSpectrumToCSV(obj, wavelength, value)
    if isempty(wavelength) || isempty(value)
        error('No valid spectrum data available to save.');
    end
    
    % Create a table with wavelength and intensity
    dataTable = table(wavelength', value', 'VariableNames', {'Wavelength_nm', 'Intensity'});

    % Generate a filename with timestamp
    fileName = ['SpectrumData_' datestr(now, 'yyyymmdd_HHMMSS') '.csv'];

    % Save the table as a CSV file
    writetable(dataTable, fileName);
    
    disp(['Spectrum data saved to: ', fileName]);
        end
    %% 

        % Destructor
        function delete(obj)    
            if ~isempty(obj.SerialCOM) && isvalid(obj.SerialCOM)
                obj.closePort();
            end
        end
    end
end