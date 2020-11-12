%% Module 3 Ankle ROM - Dorsi/Plantar Flexion
%{
ENGM 420A
Based from MATLAB
Created on 11/12/20 By Grant Wong
%}

clc; clear all; close all
%% ROM - Right Ankle
%%Setup the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 85);
% Specify sheet and range
opts.Sheet = "R_Ankle_Angle_DVJ_H";
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
minVal_R = zeros(1,numel(columns)); %Initializes variable
maxVal_R = zeros(1,numel(columns)); %Initializes variable
ROM_R = zeros(1,numel(columns)); %Initializes variable
for i = 1:columns
    sub_R{i} = data_R(1:101, i);
    minVal_R(i) = min(sub_R{i});
    maxVal_R(i) = max(sub_R{i});
    ROM_R(i) = maxVal_R(i) - minVal_R(i);
end

%% ROM - Left Ankle
%%Setup the Import Options and import the data
opts = spreadsheetImportOptions("NumVariables", 85);
% Specify sheet and range
opts.Sheet = "L_Ankle_Angle_DVJ_H";
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
minVal_L = zeros(1,numel(columns)); %Initializes variable
maxVal_L = zeros(1,numel(columns)); %Initializes variable
ROM_L = zeros(1,numel(columns)); %Initializes variable
for i = 1:columns
    sub_L{i} = data_L(1:101, i);
    minVal_L(i) = min(sub_L{i});
    maxVal_L(i) = max(sub_L{i});
    ROM_L(i) = maxVal_L(i) - minVal_L(i);
end


% data_L = readtable(xlsread('2017 Womens DVJ_H Time Series.xlsx', 'L_Ankle_Angle_DVJ_H'));
% for col = 1:2:columns
%     out{col} = data(2:101);
%     x = data(2:101, col);
%     y = data(2:101, col + 1);
%     
%     plot(x, y, '-', 'LineWidth', 2);
%     hold on;
% end
% put = data(2:101, col)