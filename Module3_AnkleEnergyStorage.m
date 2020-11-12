%% Module 3 Ankle Energy Storage
%{
ENGM 420A
Based from MATLAB
Created on 11/12/20 By Grant Wong
%}

clc; clear all; close all
%% Energy Storage - Right Ankle
%%Setup the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 85);
% Specify sheet and range
opts.Sheet = "R_Ankle_Power_BW_DVJ_H";
opts.DataRange = "B2:CH102";
% Specify column names and types
opts.VariableNames = ["S31", "S21", "S28", "S29", "S24", "S16", "S25", "S23", "S26", "S22", "S20", "S17", "S15", "S19", "S18", "S27", "S30", "S40", "S33", "S37", "S32", "S34", "S39", "S44", "S35", "S36", "S42", "S41", "S43", "S38", "S10", "S14", "S12", "S7", "S1", "S2", "S4", "S11", "S8", "S3", "S9", "S6", "S13", "S5", "S52", "S45", "S47", "S49", "S48", "S46", "S51", "S53", "S50", "S61", "S69", "S58", "S71", "S59", "S63", "S73", "S64", "S55", "S56", "S57", "S60", "S65", "S62", "S66", "S54", "S72", "S67", "S70", "S68", "S81", "S78", "S77", "S80", "S85", "S83", "S76", "S75", "S74", "S82", "S79", "S84"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
% Import the data
data_R = table2array(readtable("/Users/grantwong/Documents/MATLAB/Biomechanics/Module3_Project/2017 Womens DVJ_H Time Series.xlsx", opts, "UseExcel", false));
%%Clear temporary variables
clear opts

%%Calculation
[rows, columns] = size(data_R);
sub_R = cell(1,numel(columns)); %Initializes variable
EStorage_R = cell(1,numel(columns)); %Initializes variable
for i = 1:1%columns
    sub_R{i} = data_R(1:101, i);
    EStorage_R{i} = cumtrapz(sub_R{i});
end

%% Energy Storage - Left Ankle
%%Setup the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 85);
% Specify sheet and range
opts.Sheet = "L_Ankle_Power_BW_DVJ_H";
opts.DataRange = "B2:CH102";
% Specify column names and types
opts.VariableNames = ["S31", "S21", "S28", "S29", "S24", "S16", "S25", "S23", "S26", "S22", "S20", "S17", "S15", "S19", "S18", "S27", "S30", "S40", "S33", "S37", "S32", "S34", "S39", "S44", "S35", "S36", "S42", "S41", "S43", "S38", "S10", "S14", "S12", "S7", "S1", "S2", "S4", "S11", "S8", "S3", "S9", "S6", "S13", "S5", "S52", "S45", "S47", "S49", "S48", "S46", "S51", "S53", "S50", "S61", "S69", "S58", "S71", "S59", "S63", "S73", "S64", "S55", "S56", "S57", "S60", "S65", "S62", "S66", "S54", "S72", "S67", "S70", "S68", "S81", "S78", "S77", "S80", "S85", "S83", "S76", "S75", "S74", "S82", "S79", "S84"];
opts.VariableTypes = ["double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double", "double"];
% Import the data
data_L = table2array(readtable("/Users/grantwong/Documents/MATLAB/Biomechanics/Module3_Project/2017 Womens DVJ_H Time Series.xlsx", opts, "UseExcel", false));
%Clear temporary variables
clear opts

%%Calculation
[rows, columns] = size(data_L);
sub_L = cell(1,numel(columns)); %Initializes variable
EStorage_L = cell(1,numel(columns)); %Initializes variable
for i = 1:columns
    sub_L{i} = data_L(1:101, i);
    EStorage_L{i} = cumtrapz(sub_L{i});
end

%% Plots
figure(1)
%Raw EMG Signal
subplot(2,1,1)
plot(EStorage_R{1},'color',[0.5 0.5 0.5]);
title('Subject 1: Right Ankle Energy Storage')
xlabel('Time [s]'); ylabel('Energy Storage [Units]');
%Rectified EMG Signal with envelope
subplot(2,1,2)
plot(EStorage_L{1},'color',[0.5 0.5 0.5]);
title('Subject 1: Left Ankle Energy Storage')
xlabel('Time [s]'); ylabel('Energy Storage [Units]');