(* ::Package:: *)


% Project Goal: Use data provided in the solar energy excel file to
% formulate a script that will take a users input for TWh/y in order to
% report the specific location that will suit the users needs (i.e. an area
% that will provide the output of energy desired).
 
%Housekeeping Commands
clc
clear
close all
 
% User input for TWh/y:
Ou_D = input(' Please input a desired energy output in [TWh/y]:    ');
 
% Section to read in the data from the excel file:
 
data = xlsread('solarresourceenergy.xlsx');         
% Reads in the entire data set
 
[num,txt_Co,raw1] = xlsread('solarresourceenergy.xlsx','A:A');
% txt_Co holds the appropriate string representing a country in the excel
% file 
 
class = xlsread('solarresourceenergy.xlsx','B:B');
% Class holds an integer which represents the class in which the certain
% area is held at, 10 being the highest output and 1 being the lowest output
 
[num,txt_Cl,raw2] = xlsread('solarresourceenergy.xlsx','C:C');
% txt_Cl holds a string for the value of KWh/m^2/day, its a string rather
% than a integer due to the fact the data is given as value such as 1-2 or
%3-4
 
MWh_yr = xlsread('solarresourceenergy.xlsx','D:D');
% MWh_yr holds an integer for the given cells Megawatt hours per year
 
GW_yr = xlsread('solarresourceenergy.xlsx','E:E');
% GW_yr holds an integer for the given cells Gigawatts per year
 
TW_yr = xlsread('solarresourceenergy.xlsx','F:F');
% TW_yr holds an integer for the given cells Terawatts per year
 
KWh_yr = TW_yr.*(10^9);
% KWh_yr uses the above data vector to calculate the Killowatt hours per
% year for each cell
 
 
 
Max_TWh _yr = max(TW_yr);                        % F66
[col,row] = max(TW_yr);                         % vector location
loc = [ txt_Co(row,1)];                       % location
cla = [class(row,1)];                           % class
KW_h= [Max_TWh _yr*(10^9)];                       % KWh/yr coverage of area
pan = Ou_D*(10^9)/(cla*365*850);                % typical 1KW 1m^2 pannel will produce 850 kwh/yr class is multiplied by this number
pan_tot = [pan];                                 % note that each class of
                                                % panels has a different
                                                % cost
if Max_TWh _yr - Ou_D > 0;
    fprintf('the best location for your solar farm is a class %0 .0f area of',cla);
    disp(loc);
else 
    pan = Max_TWh _yr*(10^9)/(class(row,1)*365*850);
    pan_tot = [pan]
    Ou_D = abs(Max_TWh _yr - Ou_D);
    TW_yr(TW_yr == Max_TWh _yr) = 0;
    [col,row] = max(TW_yr);
    Max_TWh _yr = max(TW_yr);
    loc = [loc txt_Co(row,1)];
    cla = [cla class(row,1)];
        if Max_TWh _yr - Ou_D > 0
            pan = Ou_D*(10^9)/(class (row,1).*365*850);
            pan_tot = [pan pan_tot]
            fprintf('the best location for your solar farm are respectively class %0 .0f %0 .0f areas of',cla(1,1),cla(1,2));
            disp(loc);
        else
            pan = Max_TWh _yr*(10^9)/(class (row,1).*365*850);
            pan_tot = [pan pan_tot]
            Ou_D = abs(Max_TWh _yr - Ou_D);
            TW_yr(TW_yr == Max_TWh _yr) = 0;
            [col,row] = max(TW_yr);
            Max_TWh _yr = max(TW_yr);
            loc = [loc txt_Co(row,1)];
            cla = [cla class(row,1)];
                if Max_TWh _yr - Ou_D > 0
                    pan = Ou_D*(10^9)/(class (row,1).*365*850);
                    pan_tot = [pan pan_tot]
                    fprintf('the best location for your solar farm are respectively class %0 .0f %0 .0f areas of',cla(1,1),cla(1,2));
                    disp(loc);
        
                else
                    pan = Max_TWh _yr*(10^9)/(class (row,1).*365*850);
                    pan_tot = [pan pan_tot]
                    Ou_D = abs(Max_TWh _yr - Ou_D);
                    TW_yr(TW_yr == Max_TWh _yr) = 0;
                    [col,row] = max(TW_yr);
                    Max_TWh _yr = max(TW_yr);
                    loc = [loc txt_Co(row,1)];
                    cla = [cla class(row,1)];
                        if Max_TWh _yr - Ou_D > 0
                            pan = Ou_D*(10^9)/(class (row,1).*365*850);
                            pan_tot = [pan pan_tot]
                            fprintf('the best location for your solar farm are respectively class %0 .0f %0 .0f areas of',cla(1,1),cla(1,2));
                            disp(loc);
                        else
                            pan = Max_TWh _yr*(10^9)/(class (row,1).*365*850);
                            pan_tot = [pan pan_tot]
                            Ou_D = abs(Max_TWh _yr - Ou_D);
                            TW_yr(TW_yr == Max_TWh _yr) = 0;
                            [col,row] = max(TW_yr);
                            Max_TWh _yr = max(TW_yr);
                            loc = [loc txt_Co(row,1)];
                            cla = [cla class(row,1)];
                            if Max_TWh _yr - Ou_D > 0
                                pan = Ou_D*(10^9)/(class (row,1).*365*850);
                                pan_tot = [pan pan_tot]
                                fprintf('the best location for your solar farm are respectively class %0 .0f %0 .0f %0 .0f areas of',cla(1,1),cla(1,2),cla(1,3));
                                disp(loc);
                                
                                
            
                            else
                                Ou_D = abs(Max_TWh _yr - Ou_D);
                                TW_yr(TW_yr == Max_TWh _yr) = 0;
                                [col,row] = max(TW_yr);
                                Max_TWh _yr = max(TW_yr);
                                loc = [loc txt_Co(row,1)];
                                cla = [cla class(row,1)];
                                fprintf('the best location for your solar farm are respectively class %0 .0f %0 .0f %0 .0f %0 .0f areas of',cla(1,1),cla(1,2),cla(1,3),cla(1,4));
                                disp(loc);
                            end
                        end
                 end
    end
end

