%{
Tristan Spectrometer Interface (For GUI, associated with SerialCommunication) 
Last Updated: 24 March 2025
Description: MATLAB-based interface for optical spectrum analysis

Features:
- Real-time spectral measurements
- Multi-platform support (Windows/Mac/Linux)
- Automatic color identification
- Data export capabilities
- Multiple measurement overlay
%}

classdef Interface < matlab.apps.AppBase
    properties (Access = private)
        % UI Components
        UIFigure          matlab.ui.Figure
        COMPortPopup      matlab.ui.control.DropDown
        BaudRatePopup     matlab.ui.control.DropDown
        DataBitsPopup     matlab.ui.control.DropDown
        StopBitsPopup     matlab.ui.control.DropDown
        ParityPopup       matlab.ui.control.DropDown
        %HandshakePopup    matlab.ui.control.DropDown
        OpenPortButton    matlab.ui.control.Button
        ClosePortButton   matlab.ui.control.Button
        MeasureButton     matlab.ui.control.Button
        MeasureContButton matlab.ui.control.Button  %Continuously Measurement
        ClearAxesButton   matlab.ui.control.Button
        DataAxes          matlab.ui.control.UIAxes
        SaveDataButton    matlab.ui.control.Button
        
        % Serial Communication
        SerialCOM
        PlotCount
        IsMeasuringCont   logical % Flag to track continuous measurement state
    end
    %% 
    
    methods (Access = public)
        function app = Interface
            app.IsMeasuringCont = false; % Initialize continuous measurement flag
            app.initializeGUI();
        end
        %% 
        
        function delete(app)
            if ~isempty(app.SerialCOM)
                app.SerialCOM.closePort();
            end
        end
    end
    
    methods (Access = private)
        function initializeGUI(app)
            % Create and configure the figure
            app.UIFigure = uifigure('Name', 'Tristan Spectrometer Interface', ...
                'Position', [100, 100, 800, 500]);

            % All the Labels and Dropdowns
            % Label and dropdown for COM Port
            uilabel(app.UIFigure, ...
                'Text', 'Ports:', ...
                'Position', [20, 450, 60, 30], ...
                'HorizontalAlignment', 'right');
            app.COMPortPopup = uidropdown(app.UIFigure, ...
                'Position', [90, 450, 100, 30], ...
                'Items', {'Select Port'}, ...
                'Value', 'Select Port');
            
            % Label and dropdown for Baud Rate
            uilabel(app.UIFigure, ...
                'Text', 'Baud Rate:', ...
                'Position', [200, 450, 70, 30], ...
                'HorizontalAlignment', 'right');
            app.BaudRatePopup = uidropdown(app.UIFigure, ...
                'Position', [280, 450, 100, 30], ...
                'Items', {'115200', '38400', '19200', '9600', '2400'}, ...
                'Value', '115200');
            
            % Label and dropdown for Data Bits
            uilabel(app.UIFigure, ...
                'Text', 'Data Bits:', ...
                'Position', [200, 400, 70, 30], ...
                'HorizontalAlignment', 'right');
            app.DataBitsPopup = uidropdown(app.UIFigure, ...
                'Position', [280, 400, 100, 30], ...
                'Items', {'8', '7'}, ...
                'Value', '8');
            
            % Label and dropdown for Stop Bits
            uilabel(app.UIFigure, ...
                'Text', 'Stop Bits:', ...
                'Position', [200, 350, 70, 30], ...
                'HorizontalAlignment', 'right');
            app.StopBitsPopup = uidropdown(app.UIFigure, ...
                'Position', [280, 350, 100, 30], ...
                'Items', {'1', '2'}, ...
                'Value', '1');
            
            % Label and dropdown for Parity
            uilabel(app.UIFigure, ...
                'Text', 'Parity:', ...
                'Position', [200, 300, 70, 30], ...
                'HorizontalAlignment', 'right');
            app.ParityPopup = uidropdown(app.UIFigure, ...
                'Position', [280, 300, 100, 30], ...
                'Items', {'none', 'odd', 'even'}, ...
                'Value', 'none');
            
            % % Label and dropdown for Handshake
            % uilabel(app.UIFigure, ...
            %     'Text', 'Handshake:', ...
            %     'Position', [200, 250, 70, 30], ...
            %     'HorizontalAlignment', 'right');
            % app.HandshakePopup = uidropdown(app.UIFigure, ...
            %     'Position', [280, 250, 100, 30], ...
            %     'Items', {'none', 'hardware', 'software'}, ...
            %     'Value', 'none');
            
            % Buttons
            % Open Port Button
            app.OpenPortButton = uibutton(app.UIFigure, ...
                'Position', [400, 450, 100, 30], ...
                'Text', 'Open Port', ...
                'ButtonPushedFcn', @(~,~) app.openSerialPort());
            
            % Close Port Button
            app.ClosePortButton = uibutton(app.UIFigure, ...
                'Position', [520, 450, 100, 30], ...
                'Text', 'Close Port', ...
                'ButtonPushedFcn', @(~,~) app.closeSerialPort(), ...
                'Enable', 'off');
            
            % Measure Button
            app.MeasureButton = uibutton(app.UIFigure, ...
                'Position', [400, 400, 100, 30], ...
                'Text', 'Single Measure', ...
                'ButtonPushedFcn', @(~,~) app.takeMeasurement(), ...
                'Enable', 'off');

            % Measure Continuously Button
            app.MeasureContButton = uibutton(app.UIFigure, ...
                'Position', [520, 400, 150, 30], ...
                'Text', 'Measure Continuously', ...
                'ButtonPushedFcn', @(~,~) app.toggleMeasureContinuously(), ...
                'Enable', 'off');
            
            % Clear/Hold Plot Button
            app.ClearAxesButton = uibutton(app.UIFigure, ...
                'Position', [520, 360, 150, 30], ...
                'Text', 'Hold Plot', ...
                'ButtonPushedFcn', @(~,~) app.toggleHoldPlot());
            
            % Save Data as csv file Button
            app.SaveDataButton = uibutton(app.UIFigure, ...
                'Position', [520, 320, 150, 30], ...
                'Text', 'Save Data', ...
                'ButtonPushedFcn', @(~,~) app.saveSpectrumData(), ...
                'Enable', 'off');

            % Create axes for plotting
            app.DataAxes = uiaxes(app.UIFigure, ...
                'Position', [50, 50, 700, 200]);
            
            % Set axes labels
            app.DataAxes.XLabel.String = 'Wavelength (nm)';
            app.DataAxes.YLabel.String = 'Irradiance';
            app.DataAxes.Title.String = 'Spectrometer Data';
            grid(app.DataAxes, 'on');
            
            % Initialize COM port dropdown
            app.updateCOMPortList();
        end
        %% 
        
function updateCOMPortList(app)
    try
        if ispc % Windows
            % Use serialportlist for Windows
            lCOM_Port = serialportlist("available");
            if ~iscell(lCOM_Port)
                lCOM_Port = cellstr(lCOM_Port);
            end
            
        elseif ismac % macOS
            % Detect serial ports on macOS
            [status, result] = system('ls /dev/tty.* /dev/cu.*');
            if status == 0 && ~isempty(result)
                lCOM_Port = strsplit(strtrim(result));
            else
                lCOM_Port = {};
            end
            
        elseif isunix % Linux
            % Detect serial ports on Linux
            [status, result] = system('ls /dev/ttyUSB* /dev/ttyACM*');
            if status == 0 && ~isempty(result)
                lCOM_Port = strsplit(strtrim(result));
            else
                lCOM_Port = {};
            end
            
        else
            % Unsupported platform
            warning('Unsupported operating system for COM port detection');
            lCOM_Port = {};
        end

        % Update the dropdown menu
        if ~isempty(lCOM_Port)
            app.COMPortPopup.Items = lCOM_Port;
            app.COMPortPopup.Value = lCOM_Port{1};
            app.OpenPortButton.Enable = 'on';
            
            % Debug output
            disp('Available ports:');
            disp(lCOM_Port);
        else
            app.COMPortPopup.Items = {'No Ports Available'};
            app.COMPortPopup.Value = 'No Ports Available';
            app.OpenPortButton.Enable = 'off';
        end
        
    catch ME
        warning(ME.identifier,'Error detecting ports: %s', ME.message);
        app.COMPortPopup.Items = {'Detection Error'};
        app.COMPortPopup.Value = 'Detection Error';
        app.OpenPortButton.Enable = 'off';
    end
end
    
        %% 
        
        function openSerialPort(app)
            try
                % Create SerialCommunication object with selected parameters
                app.SerialCOM = SerialCommunication(...
                    app.COMPortPopup.Value, ...
                    str2double(app.BaudRatePopup.Value), ...
                    str2double(app.DataBitsPopup.Value), ...
                    str2double(app.StopBitsPopup.Value), ...
                    app.ParityPopup.Value);
                
                % Update UI state
                app.OpenPortButton.Enable = 'off';
                app.ClosePortButton.Enable = 'on';
                app.MeasureButton.Enable = 'on';
                app.MeasureContButton.Enable = 'on'; % Enable continuous measurement button
                
                % Show success message
                uialert(app.UIFigure, 'Serial port opened successfully.', 'Success','Icon', 'success');
                
            catch ME
                % Update UI state on error
                app.OpenPortButton.Enable = 'on';
                app.ClosePortButton.Enable = 'off';
                app.MeasureButton.Enable = 'off';
                app.MeasureContButton.Enable = 'off';
                uialert(app.UIFigure, ['Error opening port: ' ME.message], 'Icon', 'warning');
            end
        end
        %% 
        
        function closeSerialPort(app)
    try
        % Check if serial connection exists
        if ~isempty(app.SerialCOM)
            % Stop continuous measurement if running
                    if app.IsMeasuringCont
                        app.toggleMeasureContinuously();
                    end
            app.SerialCOM.closePort();
            app.SerialCOM = [];
            
            % Update UI state
            app.OpenPortButton.Enable = 'on';
            app.ClosePortButton.Enable = 'off';
            app.MeasureButton.Enable = 'off';
            app.MeasureContButton.Enable = 'off';
            app.SaveDataButton.Enable = 'off';  % Also disable save button
            
            % Show success message
            uialert(app.UIFigure, ...
                'Serial port closed successfully.', ...
                'Success', ...
                'Icon', 'success');
        else
            % Show warning if no port was open
            uialert(app.UIFigure, ...
                'No active serial port connection to close.', ...
                'Warning', ...
                'Icon', 'warning');
        end
        
    catch ME
        % Show error message with details
        uialert(app.UIFigure, ...
            sprintf('Failed to close port: %s', ME.message), ...
            'Error', ...
            'Icon', 'error');
        
        % Try to clean up UI state anyway
        app.OpenPortButton.Enable = 'on';
        app.ClosePortButton.Enable = 'off';
        app.MeasureButton.Enable = 'off';
        app.MeasureContButton.Enable = 'off';
        app.SaveDataButton.Enable = 'off';
    end
end
        %% 
        
        function takeMeasurement(app)
            try
                % Check if port is valid
                if isempty(app.SerialCOM)
                    error('Serial port is not connected');
                end
                
                % Get clear plot state
                clearPlot = strcmp(app.ClearAxesButton.Text, 'Hold Plot');
                
                % Restart the PlotCount 
                if clearPlot
                    app.PlotCount = 0;
                end
                
                % Increment plot counter
                app.PlotCount = app.PlotCount + 1;
                
                if app.PlotCount > 10
                    uialert(app.UIFigure, 'Cannot plot more than 10 spectra at once. Please clear the plot.', 'Warning');
                    app.PlotCount = app.PlotCount - 1;
                    return;
                end

                % Take measurement with THREE arguments
                app.SerialCOM.takeMeasurement(app.DataAxes, clearPlot, app.PlotCount);
                app.SaveDataButton.Enable = 'on';

            catch ME
                uialert(app.UIFigure, ['Measurement failed: ' ME.message], 'Error');
            end
        end
        %% 
        function toggleMeasureContinuously(app)
    try
        % Check if port is valid
        if isempty(app.SerialCOM)
            error('Serial port is not connected');
        end
        
        if ~app.IsMeasuringCont
            % Start continuous measurement
            app.IsMeasuringCont = true;
            app.MeasureContButton.Text = 'Stop Measuring';
            
            % Disable other buttons to prevent conflicts
            app.MeasureButton.Enable = 'off';
            app.SaveDataButton.Enable = 'off';
            app.ClearAxesButton.Enable = 'off';
            app.ClosePortButton.Enable = 'off';
            
            % Get clear plot state
            clearPlot = strcmp(app.ClearAxesButton.Text, 'Hold Plot');
            if clearPlot
                app.PlotCount = 0;
            end
            
            % Start continuous measurement
            app.SerialCOM.measureContinuously(app.DataAxes, clearPlot, @() app.IsMeasuringCont);
            
            % Reset state after measurement stops
            app.IsMeasuringCont = false;
            app.MeasureContButton.Text = 'Measure Continuously';
            
            % Re-enable buttons
            app.MeasureButton.Enable = 'on';
            app.SaveDataButton.Enable = 'on';
            app.ClearAxesButton.Enable = 'on';
            app.ClosePortButton.Enable = 'on';
        else
            % Stop continuous measurement
            app.IsMeasuringCont = false;
        end
        
    catch ME
        % Ensure we reset the state on error
        app.IsMeasuringCont = false;
        app.MeasureContButton.Text = 'Measure Continuously';
        app.MeasureButton.Enable = 'on';
        app.SaveDataButton.Enable = 'on';
        app.ClearAxesButton.Enable = 'on';
        app.ClosePortButton.Enable = 'on';
        uialert(app.UIFigure, ['Continuous measurement failed: ' ME.message], 'Error', 'Icon', 'error');
    end
end
        %% 
        
        function toggleHoldPlot(app)
            if strcmp(app.ClearAxesButton.Text, 'Hold Plot')
                app.ClearAxesButton.Text = 'Clear Plot';
                hold(app.DataAxes, 'on');
            else
                app.ClearAxesButton.Text = 'Hold Plot';
                hold(app.DataAxes, 'off');
                cla(app.DataAxes);
                app.PlotCount = 0;
                grid(app.DataAxes, 'on');
            end
        end
        %% 
   function saveSpectrumData(app)
    try
        % Check if serial port is connected
        if isempty(app.SerialCOM)
            error('Serial port is not connected.');
        end

        % Request spectrum data
        fprintf(app.SerialCOM.SerialCOM, 'Get spectrum cur');
        pause(2); % Allow time for data retrieval

        % Parse the spectrum data
        [wavelength, value] = app.SerialCOM.parseSpectrumData();

        % Check if data is valid before saving
        if isempty(wavelength) || isempty(value)
            error('No valid spectrum data received.');
        end

        % Save spectrum data to CSV
        app.SerialCOM.saveSpectrumToCSV(wavelength, value);

        % Show success message
        uialert(app.UIFigure, 'Spectrum data saved successfully.', 'Success');
        
    catch ME
        uialert(app.UIFigure, ['Failed to save spectrum data: ' ME.message], 'Error');
    end
   end
 end
end