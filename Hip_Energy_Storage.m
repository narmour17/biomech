%% Module 3 Hip Energy Storage (Area under power curve)
%{
ENGM 420A
Based from MATLAB
Created on 11/12/20 By Chris Stone
%}

clc; clear all; close all

%% Setup the Import Options and import the data (Left Side)
opts = spreadsheetImportOptions("NumVariables", 85);

% Specify sheet and range
opts.Sheet = "L_Hip_Power_BW_DVJ_H";
opts.DataRange = "B2:CH102";

% Specify column names and types
opts.VariableNames = ["S31", "S21", "S28", "S29", "S24", "S16", "S25", "S23", "S26", "S22", "S20", "S17", "S15", "S19", "S18", "S27", "S30", "S40", "S33", "S37", "S32", "S34", "S39", "S44", "S35", "S36", "S42", "S41", "S43", "S38", "S10", "S14", "S12", "S7", "S1", "S2", "S4", "S11", "S8", "S3", "S9", "S6", "S13", "S5", "S52", "S45", "S47", "S49", "S48", "S46", "S51", "S53", "S50", "S61", "S69", "S58", "S71", "S59", "S63", "S73", "S64", "S55", "S56", "S57", "S60", "S65", "S62", "S66", "S54", "S72", "S67", "S70", "S68", "S81", "S78", "S77", "S80", "S85", "S83", "S76", "S75", "S74", "S82", "S79", "S84"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data
dataL = table2array(readtable("\\mcse-files1\studentfiles$\cstone18\Documents\BIO 2020\Project\2017 Womens DVJ_H Time Series.xlsx", opts, "UseExcel", false));

clear opts

%% Hip Energy Storage (Left Side)

[rows, columns] = size(dataL);
for i = 1 : columns
    %Right Hip Energy 
    sub_L{i} = dataL(1:101,i);
    hipEnergy_L{i} = cumtrapz(sub_L{i});
    
end

%% Setup the Import Options and import the data (Right Side)
opts = spreadsheetImportOptions("NumVariables", 85);

% Specify sheet and range
opts.Sheet = "R_Hip_Power_BW_DVJ_H";
opts.DataRange = "B2:CH102";

% Specify column names and types
opts.VariableNames = ["S31", "S21", "S28", "S29", "S24", "S16", "S25", "S23", "S26", "S22", "S20", "S17", "S15", "S19", "S18", "S27", "S30", "S40", "S33", "S37", "S32", "S34", "S39", "S44", "S35", "S36", "S42", "S41", "S43", "S38", "S10", "S14", "S12", "S7", "S1", "S2", "S4", "S11", "S8", "S3", "S9", "S6", "S13", "S5", "S52", "S45", "S47", "S49", "S48", "S46", "S51", "S53", "S50", "S61", "S69", "S58", "S71", "S59", "S63", "S73", "S64", "S55", "S56", "S57", "S60", "S65", "S62", "S66", "S54", "S72", "S67", "S70", "S68", "S81", "S78", "S77", "S80", "S85", "S83", "S76", "S75", "S74", "S82", "S79", "S84"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];

% Import the data
dataR = table2array(readtable("\\mcse-files1\studentfiles$\cstone18\Documents\BIO 2020\Project\2017 Womens DVJ_H Time Series.xlsx", opts, "UseExcel", false));

clear opts

%% Hip Energy Storage (Left Side)

[rows, columns] = size(dataR);
for i = 1 : columns
    %Right Hip Energy 
    sub_R{i} = dataR(1:101,i);
    hipEnergy_R{i} = cumtrapz(sub_R{i});
end


% figure(1)
% %Hip Energy Storage
% subplot(2,1,1)
% plot(hipEnergy_L{i},'color',[0.5 0.5 0.5]);
% title('Hip Energy Storage Left')
% xlabel('Time [s]'); ylabel('Energy');
% %Rectified EMG Signal with envelope
% subplot(2,1,2)
% plot(hipEnergy_R{i},'color',[0.5 0.5 0.5]);
% title('Hip Energy Storage Right')
% hold on
