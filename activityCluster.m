%% Cluster Mouse Activities
% N Mikolajewicz, 03.02.19

clear all; close all;

%% User input required
file = 'allData.xlsx'; % input excel file
sheet = 'Sheet1'; % input excel sheet
kClusters = 3; % number of clusters
nReplicates = 10; % number of times clustering is performed before determining optimal solution

%% No further user input required (do not modify)

[num,txt,raw] = xlsread(file, sheet);

% assign data and check inputs
time = raw(2:end,2);

ID = num(:,1); ID = ID(~isnan(ID));
x = num(:,3); x = x(~isnan(x));
y = num(:,4); y = y(~isnan(y));
z = num(:,5); z = z(~isnan(z));

if length(time) == length(x)
else; error('check inputs'); end

% transform time so it starts at 0
dataNumber = datenum(time);
uID = unique(ID);

timeZ = nan(length(ID),1);

h1 = figure; subplot(121); 
for i = 1:length(uID);
    
    t = dataNumber(ID == uID(i));
    t = t - min(t);
    timeZ(ID == uID(i)) = t;
    
    xyz(i,1) = sum(x(ID == uID(i)));
    xyz(i,2) = sum(y(ID == uID(i)));
    xyz(i,3) = sum(z(ID == uID(i)));
    
    
  plot (1:3,  xyz(i,:)); hold on;
  leg{i} = ['ID ' num2str(uID(i))];
end

set(gca,'XTick', 1:3)
set(gca,'XTickLabel',{'x','y','z'})

xlabel('Dimension')
ylabel('Count (total)')
title('Activity Levels (input)')
legend(leg);



%% The silhouette value 
% a measure of how similar an object is to its own cluster (cohesion) compared to other clusters (separation). 
% The silhouette ranges from ?1 to +1, where a high value indicates that the object is well matched to its 
% own cluster and poorly matched to neighboring clusters. If most objects have a high value, then the 
% clustering configuration is appropriate. If many points have a low or negative value, then the clustering 
% configuration may have too many or too few clusters.

idx3 = kmeans(xyz,kClusters,'Distance','cityblock','Display','final','Replicates',nReplicates);

figure
[silh3,h] = silhouette(xyz,idx3,'cityblock');
h = gca;
h.Children.EdgeColor = [.8 .8 1];
xlabel 'Silhouette Value'
ylabel 'Cluster'
title (['Silhouette Plot (' num2str(kClusters) ' clusters)']); 
vline(0.6) %threshold value for satisfactor clustering

% assign cluster memberships to data structure
for i = 1:length(uID);
    clusterSol(i).ID = uID(i);
    clusterSol(i).membership = idx3(i);
    clusterSol(i).xTotal = xyz(i,1);
    clusterSol(i).yTotal = xyz(i,2);
    clusterSol(i).zTotal = xyz(i,3);
    
end

%% plot clusters
figure(h1)
subplot(122); 

col = {'k', 'r','b', 'g','c','m'};

for i = 1:kClusters;
    holder = [];
    holder = xyz(idx3 == i,:);
    
    results(i).Cluster = i;
    results(i).N = size(holder,1);
    results(i).xAVE = mean(holder(:,1));
    results(i).xSTD = std(holder(:,1));
    results(i).yAVE = mean(holder(:,2));
    results(i).ySTD = std(holder(:,2));
    results(i).zAVE = mean(holder(:,3));
    results(i).zSTD = std(holder(:,3));
    
    for j = 1:size(holder,1)
        plot ([1:3],  holder(j,:), col{i}); hold on;
    end
end

set(gca,'XTick', 1:3)
set(gca,'XTickLabel',{'x','y','z'})

xlabel('Dimension')
ylabel('Count (total)')
title('Activity Levels (clustered)')


figure;
for i = 1:kClusters
    if i < kClusters
        subplot(131);
        bar(i, results(i).xAVE, col{i}); hold on;
        e = errorbar(i, results(i).xAVE, results(i).xSTD/sqrt( results(i).N)); hold on;
        e.Color = col{i};
        subplot(132);
        bar(i, results(i).yAVE, col{i}); hold on; 
        e = errorbar(i, results(i).yAVE, results(i).ySTD/sqrt( results(i).N)); hold on; 
        e.Color = col{i};
        subplot(133);
        bar(i, results(i).zAVE,  col{i}); hold on; 
        e = errorbar(i, results(i).zAVE, results(i).zSTD/sqrt( results(i).N)); hold on; 
        e.Color = col{i};
    else
        subplot(131);
        bar(i, results(i).xAVE, col{i}); hold on;
        e = errorbar(i, results(i).xAVE, results(i).xSTD/sqrt( results(i).N)); hold on;
        set(gca, 'XTickLabel', ''); xlabel('Clusters'); ylabel('Count (mean)'); title('X');
        e.Color = col{i};
        subplot(132);
         bar(i, results(i).yAVE, col{i}); hold on; 
        e = errorbar(i, results(i).yAVE, results(i).ySTD/sqrt( results(i).N)); hold on; 
        set(gca, 'XTickLabel', ''); xlabel('Clusters'); ylabel('Count (mean)'); title('Y');
        e.Color = col{i};
        subplot(133);
         bar(i, results(i).zAVE,  col{i}); hold on; 
        e = errorbar(i, results(i).zAVE, results(i).zSTD/sqrt( results(i).N)); hold on;
        set(gca, 'XTickLabel', ''); xlabel('Clusters'); ylabel('Count (mean)'); title('Z');  
        e.Color = col{i};
    end 
end


try;
    now = datestr(datetime('now'));
    now = strrep(now, ':','-');
    
    file = strrep(file, '.xlsx',' ');
    file = [file 'RESULTS ' now '.xlsx'];
% nowT = now;
writetable(struct2table(clusterSol), file, 'Sheet' ,['clusters']);
writetable(struct2table(results), file, 'Sheet' ,['results']);
catch;
    error(' ')
end
RemoveSheet123(file);
display('Clustering Complete!'); 

function hhh=vline(x,in1,in2)
% function h=vline(x, linetype, label)
% 
% Draws a vertical line on the current axes at the location specified by 'x'.  Optional arguments are
% 'linetype' (default is 'r:') and 'label', which applies a text label to the graph near the line.  The
% label appears in the same color as the line.
%
% The line is held on the current axes, and after plotting the line, the function returns the axes to
% its prior hold state.
%
% The HandleVisibility property of the line object is set to "off", so not only does it not appear on
% legends, but it is not findable by using findobj.  Specifying an output argument causes the function to
% return a handle to the line, so it can be manipulated or deleted.  Also, the HandleVisibility can be 
% overridden by setting the root's ShowHiddenHandles property to on.
%
% h = vline(42,'g','The Answer')
%
% returns a handle to a green vertical line on the current axes at x=42, and creates a text object on
% the current axes, close to the line, which reads "The Answer".
%
% vline also supports vector inputs to draw multiple lines at once.  For example,
%
% vline([4 8 12],{'g','r','b'},{'l1','lab2','LABELC'})
%
% draws three lines with the appropriate labels and colors.
% 
% By Brandon Kuczenski for Kensington Labs.
% brandon_kuczenski@kensingtonlabs.com
% 8 November 2001
if length(x)>1  % vector input
    for I=1:length(x)
        switch nargin
        case 1
            linetype='r:';
            label='';
        case 2
            if ~iscell(in1)
                in1={in1};
            end
            if I>length(in1)
                linetype=in1{end};
            else
                linetype=in1{I};
            end
            label='';
        case 3
            if ~iscell(in1)
                in1={in1};
            end
            if ~iscell(in2)
                in2={in2};
            end
            if I>length(in1)
                linetype=in1{end};
            else
                linetype=in1{I};
            end
            if I>length(in2)
                label=in2{end};
            else
                label=in2{I};
            end
        end
        h(I)=vline(x(I),linetype,label);
    end
else
    switch nargin
    case 1
        linetype='r:';
        label='';
    case 2
        linetype=in1;
        label='';
    case 3
        linetype=in1;
        label=in2;
    end
    
    
    
    g=ishold(gca);
    hold on
    y=get(gca,'ylim');
    h=plot([x x],y,linetype);
    if length(label)
        xx=get(gca,'xlim');
        xrange=xx(2)-xx(1);
        xunit=(x-xx(1))/xrange;
        if xunit<0.8
            text(x+0.01*xrange,y(1)+0.1*(y(2)-y(1)),label,'color',get(h,'color'))
        else
            text(x-.05*xrange,y(1)+0.1*(y(2)-y(1)),label,'color',get(h,'color'))
        end
    end     
    if g==0
    hold off
    end
    set(h,'tag','vline','handlevisibility','off')
end % else
if nargout
    hhh=h;
end
end

function RemoveSheet123(excelFileName,sheetName)
% RemoveSheet123 - removes the sheets that are automatically added to excel
% file. 
% When Matlab writes data to a new Excel file, the Excel software
% automatically creates 3 sheets (the names are depended on the user
% languade). This appears even if the user has defined the sheet name to be
% added. 
%
% Usage:
% RemoveSheet123(excelFileName) - remove "sheet1", "sheet2","sheet3" from
% the excel file. excelFileName is a string of the Excel file name.
% RemoveSheet123(excelFileName,sheetName) - enables the user to enter the
% sheet name when the language is other than English.
% sheetName is the default sheet name, without the number.
%
%
%                       Written by Noam Greenboim
%                       www.perigee.co.il
%
%% check input arguments
if nargin < 1 || isempty(excelFileName)
    error('Filename must be specified.');
end
if ~ischar(excelFileName)
    error('Filename must be a string.');
end
try
    excelFileName = validpath(excelFileName);
catch 
    error('File not found.');
end
if nargin < 2
    sheetName = 'Sheet'; % EN: Sheet, DE: Tabelle, HE: ?????? , etc. (Lang. dependent)
else
    if ~ischar(sheetName)
        error('Default sheet name must be a string.');
    end
end
%%
% Open Excel file.
objExcel = actxserver('Excel.Application');
objExcel.Workbooks.Open(excelFileName); % Full path is necessary!
% Delete sheets.
try
      % Throws an error if the sheets do not exist.
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '1']).Delete;
      fprintf('\nsheet #1 - deleted.')
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '2']).Delete;
      fprintf('\nsheet #2 - deleted.')
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '3']).Delete;
      fprintf('\nsheet #3 - deleted.\n')
catch
    fprintf('\n')
    O=objExcel.ActiveWorkbook.Worksheets.get;
    if O.Count==1
        error('Can''t delete the last sheet. Excel file must containt at least one sheet.')
    else
      warning('Problem occured. Check excel file.'); 
    end
end
% Save, close and clean up.
objExcel.ActiveWorkbook.Save;
objExcel.ActiveWorkbook.Close;
objExcel.Quit;
objExcel.delete;
end
function filenameOut = validpath(filename)
    % VALIDPATH builds a full path from a partial path specification
    %   FILENAME = VALIDPATH(FILENAME) returns a string vector containing full
    %   path to a file. FILENAME is string vector containing a partial path
    %   ending in a file or directory name. May contain ..\  or ../ or \\. The
    %   current directory (pwd) is prepended to create a full path if
    %   necessary. On UNIX, when the path starts with a tilde, '~', then the
    %   current directory is not prepended.
    %
    %   See also XLSREAD, XLSWRITE, XLSFINFO.
    
    %   Copyright 1984-2012 The MathWorks, Inc.
    
    %First check for wild cards, since that is not supported.
    if strfind(filename, '*') > 0
        error(message('MATLAB:xlsread:Wildcard', filename));
    end
    
    % break partial path in to file path parts.
    [Directory, file, ext] = fileparts(filename);
    if ~isempty(ext)
        filenameOut = getFullName(filename);
    else
        extIn = matlab.io.internal.xlsreadSupportedExtensions;
        for i=1:length(extIn)
            try                                                                %#ok<TRYNC>
                filenameOut = getFullName(fullfile(Directory, [file, extIn{i}]));
                return;
            end
        end
        error(message('MATLAB:xlsread:FileDoesNotExist', filename));    
    end
end
function absolutepath=abspath(partialpath)
    
    % parse partial path into path parts
    [pathname, filename, ext] = fileparts(partialpath);
    % no path qualification is present in partial path; assume parent is pwd, except
    % when path string starts with '~' or is identical to '~'.
    if isempty(pathname) && strncmp('~', partialpath, 1)
        Directory = pwd;
    elseif isempty(regexp(partialpath,'(.:|\\\\)', 'once')) && ...
            ~strncmp('/', partialpath, 1) && ...
            ~strncmp('~', partialpath, 1);
        % path did not start with any of drive name, UNC path or '~'.
        Directory = [pwd,filesep,pathname];
    else
        % path content present in partial path; assume relative to current directory,
        % or absolute.
        Directory = pathname;
    end
    
    % construct absolute filename
    absolutepath = fullfile(Directory,[filename,ext]);
end
function filename = getFullName(filename)
    FileOnPath = which(filename);
    if isempty(FileOnPath)
        % construct full path to source file
        filename = abspath(filename);
        if isempty(dir(filename)) && ~isdir(filename)
            % file does not exist. Terminate importation of file.
            error(message('MATLAB:xlsread:FileDoesNotExist', filename));
        end
    else
        filename = FileOnPath;
    end
end
