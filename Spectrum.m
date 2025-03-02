%{
Spectrum (For execution on Command Window)
Last Updated: 27 February 2025
Description: MATLAB-based interface for optical spectrum analysis
Features:
- Real-time spectral measurements
- Multi-platform support (Windows/Mac)
- Automatic color identification
- Data export capabilities
- Multiple measurement overlay
%}

classdef Spectrum
    properties (Access = public)
        SerialCOM  % Serial port object
    end
    %% 
    
    methods (Access = public)
        % Constructor
        function obj = Spectrum(port)
            disp('Initializing spectrometer...');
            availablePorts = serialportlist("available");

            if isempty(availablePorts)
                error('No available serial ports detected.');
            end

            disp('Available Ports:');
            for i = 1:length(availablePorts)
                disp([num2str(i) '. ' availablePorts{i}]);
            end

            disp('0. Exit');

            while true
                portIndex = str2double(input('Enter port number (or 0 to exit): ', 's'));
                if isnan(portIndex) || portIndex < 0 || portIndex > length(availablePorts)
                    disp(['Enter a number between 0 and ', num2str(length(availablePorts))]);
                    continue;
                elseif portIndex == 0
                    disp('Exiting...');
                    obj.SerialCOM = [];  % Initialize SerialCOM to empty
                    return;
                end
                break;
            end

            selectedPort = availablePorts{portIndex};
            obj.SerialCOM = serialport(selectedPort, 115200, 'Timeout', 10);
            configureTerminator(obj.SerialCOM, 'LF');
            disp(['Connected to ' selectedPort]);
            obj.runInteractive();
        end
%% 

        % Close the serial port
        function closePort(obj)
            if ~isempty(obj.SerialCOM) && isvalid(obj.SerialCOM)
                delete(obj.SerialCOM);  % Properly close the serial port
                disp('Serial port closed.');
            end
        end

%% 
        % Run interactive menu
        function runInteractive(obj)
            measurements = {};  % Stores multiple spectra
            PlotCount = 0;  % Track number of plots

            while true
                disp('Functions:');
                menu = {'New Single Measurement', 'Multiple Measurements', 'Determine Color', 'Save Data as CSV', 'Close Port'};
                % if ~isempty(measurements)
                %     menu = {'New Single Measurement', 'Multiple Measurements', ...
                %            'Determine Color', 'Save Data as CSV', 'Close Port'};
                % end

                for i = 1:length(menu)
                    disp([num2str(i) '. ' menu{i}]);
                end

                choice = str2double(input('Enter choice: ', 's'));
                if isnan(choice) || choice < 1 || choice > length(menu)
                    disp(['Enter a number between 1 and ' num2str(length(menu))]);
                    continue;
                end

                switch choice
                    case 1  % Single Measurement
                        [wavelength, value] = obj.takeMeasurement();  % Get measurement data

                        if ~isempty(wavelength) && ~isempty(value) && length(wavelength) == length(value)
                            PlotCount = 1;
                            figure(1);  % Create figure for plotting
                            clf;  % Clear the figure initially
                            hold on;  % Hold the current plot

                            obj.plotSpectrum(wavelength, value, PlotCount);  % Plot the new spectrum

                            % Store the first measurement in the cell array
                            measurements = {{wavelength, value}};

                            % % Store data in the base workspace
                            % obj.storeDataInWorkspace(wavelength, value);
                            % pause(0.1);
                        else
                            disp('Warning: Measurement failed or returned invalid data.');
                        end

                    case 2  % Multiple Measurements
                        [newWavelength, newValue] = obj.takeMeasurement();  % Get new measurement data

                        if ~isempty(newWavelength) && ~isempty(newValue) && length(newWavelength) == length(newValue)
                             
                            if ~exist('measurements', 'var') || isempty(measurements)
                            % First measurement
                            PlotCount = 1;
                            figure(1); % Create figure for plotting
                            clf; % Clear the figure initially
                            hold on; % Hold the current plot
        
                            % Initialize measurements cell array with the first measurement
                            measurements = {{newWavelength, newValue}};
                            else
                            % Subsequent measurements
                            PlotCount = PlotCount + 1;
        
                            % Store the new measurement in the measurements cell array
                            measurements{end + 1} = {newWavelength, newValue};
                            end
                            
                            % Plot only the new spectrum without clearing the entire figure
                            obj.plotSpectrum(newWavelength, newValue, PlotCount); 

                            % % Store data in the base workspace
                            % obj.storeDataInWorkspace(newWavelength, newValue);
                            % pause(0.1);
                        else
                            disp('Warning: Measurement failed or returned invalid data.');
                        end

                    case 3  % Determine Color
                        if isempty(measurements)
                            disp('Warning: No measurement data available.');
                        else
                            latestMeasurement = measurements{end};
                            obj.determineColor(latestMeasurement{1}, latestMeasurement{2});
                        end

                    case 4  % Save Data as CSV
                        if isempty(measurements)
                            disp('Warning: No measurement data available.');
                        else
                            if numel(measurements) > 1
                                while true
                                    saveChoice = input('Save recent (1) or all (2)? ');
                                    if ismember(saveChoice, [1, 2])
                                        break;
                                    end
                                    disp('Enter 1 or 2.');
                                end
                                if saveChoice == 2
                                    obj.saveMultipleSpectraToCSV(measurements);
                                else
                                    latestMeasurement = measurements{end};
                                    obj.saveSpectrumToCSV(latestMeasurement{1}, latestMeasurement{2});
                                end
                            else
                                latestMeasurement = measurements{1};
                                obj.saveSpectrumToCSV(latestMeasurement{1}, latestMeasurement{2});
                            end
                        end

                    case 5  % Close Port
                        obj.closePort();
                        return;
                end
            end
        end
%% 

        % Take a single measurement
        function [wavelength, value] = takeMeasurement(obj)
            disp('Starting measurement...');
            flush(obj.SerialCOM);  %clear the serial buffer
            fprintf(obj.SerialCOM, 'Start single');
            disp('Measuring...');
            pause(5);

            fprintf(obj.SerialCOM, 'Get spectrum cur');
            disp('Getting spectrum...');
            pause(2);

            [wavelength, value] = obj.parseSpectrumData();

            if isempty(wavelength) || isempty(value)
                disp('Warning: Measurement failed or returned invalid data. Check the connection and try again.');
            else
                disp('Measurement complete.');
            end
        end
%% 

        % Parse spectrum data from the serial port
        function [Wavelength, Value] = parseSpectrumData(obj)
            Wavelength = [];
            Value = [];
            disp('Parsing data...');
            timeout = tic;
            while obj.SerialCOM.NumBytesAvailable > 0 && toc(timeout) < 5
                line = readline(obj.SerialCOM);
                data = sscanf(line, '%f %f');
                if numel(data) == 2
                    Wavelength(end+1) = data(1);
                    Value(end+1) = data(2);
                end
            end
            if isempty(Wavelength) || isempty(Value)
                disp('Warning: No data received from the spectrometer.');
            else
                disp('Parsing complete.');
            end
        end
%% 

        % Plot the spectrum
        function plotSpectrum(obj, wavelength, value, PlotCount)
            figure(1);
            hold on;

            if ~isempty(wavelength) && ~isempty(value)
                wavelength(end+1) = NaN;
                value(end+1) = NaN;

                % Normalize the intensity values by dividing by the maximum value
                normalizedValue = value / max(value);

                % Find the wavelength with maximum intensity (peak)
                colorName = obj.getColorDescription(wavelength, normalizedValue);
                colorRGB = obj.getColorRGB(colorName);

                % Plot the spectrum with normalized y-axis
                plot(wavelength, normalizedValue, 'Color', colorRGB, 'LineWidth', 2, ...
                     'DisplayName', ['Line ' num2str(PlotCount) ' (' colorName ')']);

                grid on;
                xlabel('Wavelength (nm)');
                ylabel('Relative Power of LED');

                % Enable legend with dynamic names
                legend('show', 'Location', 'best');

                % Set y-axis limits to [0, 1] for relative power
                ylim([0, 1]);

                % Customize y-axis ticks and labels
                yticks([0, 0.2, 0.4, 0.6, 0.8, 1.0]);
                yticklabels({'0', '0.2', '0.4', '0.6', '0.8', '1.0'});

                xlim([min(wavelength) max(wavelength)]);
                title('Spectrometer Data');
                hold off;
            else
                disp('Warning: No wavelength or intensity data to plot.');
            end
        end
%% 

        % Determine the color of the light
        function color = determineColor(obj, wavelength, value)
            if isempty(wavelength) || isempty(value)
                disp('Warning: No data available.');
                color = 'Unknown';
            else
                color = obj.getColorDescription(wavelength, value);
                disp(['Light color: ', color]);
            end
        end
%% 

        % Save a single spectrum to CSV
        function saveSpectrumToCSV(obj, wavelength, value)
            filename = ['SpectrumData_' datestr(now,'yyyymmdd_HHMMSS') '.csv'];
            dataTable = table(wavelength', value', 'VariableNames', {'Wavelength_nm', 'Intensity'});
            writetable(dataTable, filename);
            disp(['Saved to: ', filename]);
        end
%% 

        % Save multiple spectra to CSV
        function saveMultipleSpectraToCSV(obj, measurements)
            filename = ['MultiSpectrum_' datestr(now, 'yyyymmdd_HHMMSS') '.csv'];

            % Find the maximum length among all measurements
            maxRows = max(cellfun(@(m) length(m{1}), measurements));

            % Initialize an empty table
            allData = table();

            for i = 1:length(measurements)
                wavelength = measurements{i}{1}';
                value = measurements{i}{2}';

                % Pad shorter columns with NaN
                if length(wavelength) < maxRows
                    wavelength(end+1:maxRows, 1) = NaN;
                    value(end+1:maxRows, 1) = NaN;
                end

                % Add to table with unique column names
                allData = [allData, table(wavelength, value, ...
                          'VariableNames', {['Wavelength_' num2str(i)], ['Intensity_' num2str(i)]})];
            end

            % Save to CSV
            writetable(allData, filename);
            disp(['Saved all measurements to: ', filename]);
        end
%% 

        % Get color description based on wavelength and intensity
        function colorName = getColorDescription(obj, wavelength, intensity)
            try
                % First smooth the intensity data to reduce noise
                windowSize = 15;
                smoothedIntensity = movmean(intensity, windowSize);

                % Define spectral ranges for different colors
                validRange = wavelength >= 250 & wavelength <= 750;

                % Work with smoothed data in valid range
                validIntensities = smoothedIntensity(validRange);
                validWavelengths = wavelength(validRange);

                % Find the maximum intensity and its wavelength
                [maxIntensity, maxIndex] = max(validIntensities);
                peakWavelength = validWavelengths(maxIndex);

                % Check for the presence of characteristic peaks for white LED
                bluePeakRange = validWavelengths >= 440 & validWavelengths <= 470;  % Blue peak range (450- 460 nm)
                yellowPeakRange = validWavelengths >= 520 & validWavelengths <= 560;  % Green-yellow peak range (530 nm)

                % Find the intensity values in these ranges
                bluePeakIntensity = max(validIntensities(bluePeakRange));
                yellowPeakIntensity = max(validIntensities(yellowPeakRange));

                % Define a threshold for recognizing a white LED
                whiteLEDThreshold = 0.3;  % Relative intensity threshold for secondary peaks

                % Check if both peaks are detected and meet the threshold for classification as a white LED
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

        % Get RGB color for plotting
        function colorRGB = getColorRGB(obj, colorName)
            switch lower(colorName)
                case 'uv/purple'
                    colorRGB = [0.5, 0, 1];
                case 'blue'
                    colorRGB = [0, 0, 1];
                case 'green'
                    colorRGB = [0, 1, 0];
                case 'yellow'
                    colorRGB = [1, 1, 0];
                case 'orange'
                    colorRGB = [1, 0.5, 0];
                case 'red'
                    colorRGB = [1, 0, 0];
                case 'white led'
                    colorRGB = [0, 0, 0];  % Black color for white LED
                otherwise
                    colorRGB = [0.5, 0.5, 0.5];  % Default gray for unknown colors
            end
        end
%% 

    %Store data in the base workspace
    %function storeDataInWorkspace(obj, wavelength, value)
    % assignin('base', 'wavelength', wavelength);
    % assignin('base', 'value', value);
    % disp('Data stored in workspace as "wavelength" and "value".');
    % drawnow;
    % end
    end
end