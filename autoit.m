classdef autoit
    % AUTOIT 
    %  AUTOIT automation interface
    %
    %   Author(s):  Alessandro RUGGERI
    %
    %   Rev 1.0-25/04/2021: first issue

%% ########### PRIVATE CONSTANTS ###############################################
    properties (Constant, Access = private)
        LIBALIAS = 'AutoItX3';
        CLIP_SIZE = 1000;
        CLIP_STR=zeros(1,autoit.CLIP_SIZE,'uint16');
        DEFAULT_WIN_TIMEOUT=5;
    end

%% ########### METHODS #########################################################
    methods
        function obj=autoit()
            error('You cannot create an istance of class autoit');
        end
    end

%% ########### STATIC METHODS ##################################################
    methods (Static)
%% --- Mouse related methods ---------------------------------------------------
        function [pt_x,y]=mouseGetPos()
            pt=calllib(autoit.LIBALIAS, 'AU3_MouseGetPos', []);
            if nargout<2
                pt_x=pt;
            elseif nargout==2
                pt_x=pt.x;
                y=pt.y;
            end
        end

        function mouseMove(x, y, speed)
            autoit.validateInputNum(x,-inf,inf);
            autoit.validateInputNum(y,-inf,inf);
            if nargin < 3
                speed=10;
            else
                speed=autoit.parseSpeed(speed);
            end
            err=calllib(autoit.LIBALIAS, 'AU3_MouseMove', x, y, speed);
            autoit.checkError(err);
        end

        function mouseClick(button, x, y, nClics, speed)
            btt=autoit.parseMouseBtt(button);
            autoit.validateInputNum(x,-inf,inf);
            autoit.validateInputNum(y,-inf,inf);
            if nargin < 4
                nClics=1;
            else
                autoit.validateInputNum(nClics,0,inf);
            end
            if nargin < 5
                speed=10;
            else
                speed=autoit.parseSpeed(speed);
            end
            err=calllib(autoit.LIBALIAS, 'AU3_MouseClick', ...
                btt, x, y, nClics, speed);
            autoit.checkError(err);
        end

%% --- Send commands methods ---------------------------------------------------
        function send(text,flag)
            if nargin < 2
                flag=0;
            end
            u16str=autoit.validateStr(text);
            flag=autoit.validateBool(flag);
            calllib(autoit.LIBALIAS, 'AU3_Send', u16str, flag);
        end

%% --- Clipboard related methods -----------------------------------------------
        function clip=clipGet()
            voidStr=libpointer('uint16Ptr',autoit.CLIP_STR);
            calllib(autoit.LIBALIAS, 'AU3_ClipGet', voidStr, autoit.CLIP_SIZE);
            idx=find(voidStr.Value==0,1);
            clip=char(voidStr.Value(1:idx-1));
        end

        function clipPut(str)
            u16str=autoit.validateStr(str);
            calllib(autoit.LIBALIAS, 'AU3_ClipPut', u16str);
        end

%% --- control related methods -------------------------------------------------
        function controlClick(title, text, controlID, button, nClicks, x, y)
            if nargin < 6
                x=0;
                y=0;
            end
            if nargin < 5
                nClicks=1;
            end
            if nargin < 4
                button='left';
            end
                
            u16title=autoit.validateStr(title);
            u16text=autoit.validateStr(text);
            u16controlID=autoit.validateStr(controlID);
            btt=autoit.parseMouseBtt(button);
            err=calllib(autoit.LIBALIAS, 'AU3_ControlClick', u16title, ...
                u16text, u16controlID, btt, nClicks, x, y);
            autoit.checkError(err);
        end
        
%% --- window related methods --------------------------------------------------
        function winActivate(title, text)
            if nargin < 2
                text='';
            end
            u16title=autoit.validateStr(title);
            u16text=autoit.validateStr(text);
            calllib(autoit.LIBALIAS, 'AU3_WinActivate', u16title, u16text);
        end

        function winWaitActive(title, text, timeout)
            if nargin < 2
                text='';
                timeout=autoit.DEFAULT_WIN_TIMEOUT;
            elseif nargin < 3
                timeout=autoit.DEFAULT_WIN_TIMEOUT;
            end
            u16title=autoit.validateStr(title);
            u16text=autoit.validateStr(text);
            autoit.validateInputNum(timeout,0,inf);
            calllib(autoit.LIBALIAS, 'AU3_WinWaitActive', ...
                u16title, u16text, timeout);
        end

        function winWaitActivate(title, text, timeout)
            if nargin < 2
                text='';
                timeout=autoit.DEFAULT_WIN_TIMEOUT;
            elseif nargin < 3
                timeout=autoit.DEFAULT_WIN_TIMEOUT;
            end
            autoit.winActivate(title, text);
            autoit.winWaitActive(title, text, timeout);
        end
        
        function winMove(title, text, x, y, width, height)
            u16title=autoit.validateStr(title);
            u16text=autoit.validateStr(text);
            autoit.validateInputNum(x,-inf,inf);
            autoit.validateInputNum(y,-inf,inf);
            autoit.validateInputNum(width,-inf,inf);
            autoit.validateInputNum(height,-inf,inf);
            calllib(autoit.LIBALIAS, 'AU3_WinMove', ...
                u16title, u16text, x, y, width, height);
        end

%% --- Other methods -----------------------------------------------------------
        function old_val=setOption(option, value)
            u16str=autoit.validateStr(option);
            autoit.validateInputNum(value,0,inf);
            ret=calllib(autoit.LIBALIAS, 'AU3_AutoItSetOption', u16str, value);
            if nargout>=1
                old_val=ret;
            end
        end

        function tf=isAdmin()
            tf=(calllib(autoit.LIBALIAS, 'AU3_IsAdmin')==1);
        end

%% --- library handling methods ------------------------------------------------
        function yes = libIsLoaded()
            yes = libisloaded(autoit.LIBALIAS);
        end

        function mex_out=libLoad()
            headerfname = 'AutoItX3_DLL.h';
            if ~autoit.libIsLoaded()
                switch(computer('arch'))
                    case 'win32'
                        dllfname = 'AutoItX3.dll';
                    case 'win64'
                        dllfname = 'AutoItX3_x64.dll';
                    otherwise
                        autoit.printError('wrongOS',...
                            'Not supported operating system');
                end
                if exist('autoit_header.m', 'file')
                    loadlibrary(dllfname, @autoit_header, ...
                        'alias', autoit.LIBALIAS);
                else
                    [~,mex]=loadlibrary(dllfname, headerfname, ...
                        'alias', autoit.LIBALIAS, ...
                        'mfilename', 'autoit_header.m');
                    if nargout >= 1
                        mex_out=mex;
                    end
                end
                if ~autoit.libIsLoaded()
                   autoit.printError('loadFailed',...
                       'Unable to load the autoit library');
                end
            end
        end

        function libUnload()
            if autoit.libIsLoaded()
                unloadlibrary(autoit.LIBALIAS);
                if libisloaded(autoit.LIBALIAS)
                    autoit.printError('unloadFailed',...
                        'Unable to unload the autoit library');
                end
            end 
        end
    end

%% ########### STATIC + PRIVATE METHODS ########################################
    methods (Static, Access = private)
%% --- Parse/Validate methods --------------------------------------------------
        function out=parseSpeed(speed)
            if ischar(speed)
                switch speed
                    case 'immediate',   out=0;
                    case 'fastest',     out=1;
                    case 'fast',        out=2;
                    case 'normal',      out=10;
                    case 'slow',        out=50;
                    case 'slowest',     out=100;
                    otherwise, autoit.printError('badInput',...
                            ['Unknown option ''' speed ''''])
                end
            else
                autoit.validateInputNum(speed,0,100);
                out=speed;
            end
        end

        function str=parseMouseBtt(btt)
            if ischar(btt)
                switch btt
                    case {'left', 'right', 'middle', 'main', 'menu', ...
                            'primary', 'secondary'}, str=[uint16(btt) 0];
                    otherwise, autoit.printError('badInput',...
                            ['Unknown option ''' btt ''''])
                end
            end
        end

        function u16str=validateStr(str)
            if ~ischar(str)
                autoit.printError('badInput','Input value is not a string');
            else
                u16str=[uint16(str) 0];
            end
        end

        function validateInputNum(val, min, max)
            if isnumeric(val) && numel(val) == 1
                if val < min || val > max
                    autoit.printError('badInput',['Input value ' ...
                        num2str(val) ' must be between ' ...
                        num2str(min) ' and ' num2str(max)]);
                end
            else
                autoit.printError('badInput',...
                    'Input value must be a scalar number')
            end
        end

        function val=validateBool(val)
            if (isnumeric(val) || islogical(val)) && numel(val) == 1
                val=double(val);
            else
                autoit.printError('badInput',...
                    'Input value must be a scalar number or boolean')
            end
        end

%% --- Error handling methods --------------------------------------------------
        function checkError(err)
            if err~=1
                ME = MException(['autoit:' num2str(err)], num2str(err));
                throwAsCaller(ME);
            end
        end

        function printError(type, msg)
            ME = MException(['autoit:' type], msg);
            throwAsCaller(ME);
        end
    end
end
