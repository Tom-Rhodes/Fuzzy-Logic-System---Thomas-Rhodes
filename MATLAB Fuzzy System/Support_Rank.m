clc
filename = ('Support_Rank.xls');

%******************************************

%   Mitigation Effectiveness Calculation

%******************************************
mitigationData = xlsread(filename, 1);
fprintf('Mitigation/Total Damage = Mitigation percentage relative to Total Damage taken\n\n');
for i=1:size(mitigationData,1)
    mitigationEffect = ((mitigationData(i,1))/(mitigationData(i,2)))*100;
    xlswrite(filename, mitigationEffect, 1 , sprintf('C%d',i+1));
    xlswrite(filename, mitigationEffect, 2 , sprintf('B%d',i+1));
    fprintf('%.2f / %.2f = %.2f\n',mitigationData(i, 1),mitigationData(i,2),mitigationEffect);        
end

%##########################################################################

%******************

%   Tank Machine

%******************
tankData = xlsread(filename, 2);
a = newfis('Tank');

a = addvar(a, 'input', 'CC Score', [0 75]);

a = addmf(a, 'input',1, 'Very low', 'trimf', [-5 0 7.5]);
a = addmf(a, 'input',1, 'Low', 'gaussmf', [7.5 18.75]);
a = addmf(a, 'input',1, 'Average', 'gaussmf', [7.5 37.5]);
a = addmf(a, 'input',1, 'High', 'gaussmf', [7.5 56.25]);
a = addmf(a, 'input',1, 'Very high', 'trimf', [60 75 80]);


a = addvar(a, 'input', 'Mitigation Effectiveness', [0 300]);

a = addmf(a, 'input',2, 'Very low', 'trimf', [-5 0 30]);
a = addmf(a, 'input',2, 'Low', 'gaussmf', [30 75]);
a = addmf(a, 'input',2, 'Average', 'gaussmf', [30 150]);
a = addmf(a, 'input',2, 'High', 'gaussmf', [30 225]);
a = addmf(a, 'input',2, 'Very high', 'trimf', [240 300 305]);

a = addvar(a,'output', 'Tank Ability', [0 100]);

a = addmf(a, 'output',1, 'Very poor', 'gaussmf', [10 0]);
a = addmf(a, 'output',1, 'Poor', 'gaussmf', [10 25]);
a = addmf(a, 'output',1, 'Standard', 'gaussmf', [10 50]);
a = addmf(a, 'output',1, 'Good', 'gaussmf', [10 75]);
a = addmf(a, 'output',1, 'Very good', 'gaussmf', [10 100]);


rule1 = [1 1 1 1 1]; %Very low CC
rule2 = [1 2 1 1 1];
rule3 = [1 3 2 1 1];
rule4 = [1 4 3 1 1];
rule5 = [1 5 3 1 1];

rule6 = [2 1 1 1 1]; %Low CC
rule7 = [2 2 2 1 1];
rule8 = [2 3 3 1 1];
rule9 = [2 4 3 1 1];
rule10 = [2 5 4 1 1];

rule11 = [3 1 2 1 1]; %Average CC
rule12 = [3 2 3 1 1];
rule13 = [3 3 3 1 1];
rule14 = [3 4 4 1 1];
rule15 = [3 5 4 1 1];

rule16 = [4 1 3 1 1]; %High CC
rule17 = [4 2 3 1 1];
rule18 = [4 3 4 1 1];
rule19 = [4 4 4 1 1];
rule20 = [4 5 5 1 1];

rule21 = [5 1 3 1 1]; %Very high CC
rule22 = [5 2 4 1 1];
rule23 = [5 3 4 1 1];
rule24 = [5 4 5 1 1];
rule25 = [5 5 5 1 1];

ruleListA = [rule1; rule2; rule3; rule4;
    rule5; rule6; rule7; rule8; rule9;
    rule10; rule11; rule12; rule13;
    rule14; rule15; rule16; rule17;
    rule18; rule19; rule20; rule21;
    rule22; rule23; rule24; rule25;];


a = addrule(a,ruleListA);

fprintf('\nTank machine:\n\n')


rule = showrule(a)

%defuzzification methods
a.defuzzMethod = 'centroid';

for i=1:size(tankData,1)
        evalplayer = evalfis([tankData(i, 1), tankData(i, 2)], a);
        fprintf('%d) In(1): %.2f, In(2) %.2f,  => Out: %.2f \n',i,tankData(i, 1),tankData(i, 2), evalplayer);  
        xlswrite(filename, evalplayer, 2 , sprintf('C%d',i+1));
        xlswrite(filename, evalplayer, 5 , sprintf('B%d',i+1));
end
%##########################################################################

%********************

%   Healer Machine

%********************
healerData = xlsread(filename, 3);
b = newfis('Healer');

b = addvar(b, 'input', 'CC Score', [0 75]);

b = addmf(b, 'input',1, 'Very low', 'trimf', [-5 0 7.5]);
b = addmf(b, 'input',1, 'Low', 'gaussmf', [7.5 18.75]);
b = addmf(b, 'input',1, 'Average', 'gaussmf', [7.5 37.5]);
b = addmf(b, 'input',1, 'High', 'gaussmf', [7.5 56.25]);
b = addmf(b, 'input',1, 'Very high', 'trimf', [60 75 80]);


b = addvar(b, 'input', 'Total Healing', [0 70000]);

b = addmf(b, 'input',2, 'Very low', 'trimf', [-5 0 7000]);
b = addmf(b, 'input',2, 'Low', 'gaussmf', [7000 17500]);
b = addmf(b, 'input',2, 'Average', 'gaussmf', [7000 35000]);
b = addmf(b, 'input',2, 'High', 'gaussmf', [7000 52500]);
b = addmf(b, 'input',2, 'Very high', 'trimf', [56000 70000 70005]);


b = addvar(b,'output', 'Healer Ability', [0 100]);

b = addmf(b, 'output',1, 'Very poor', 'gaussmf', [10 0]);
b = addmf(b, 'output',1, 'Poor', 'gaussmf', [10 25]);
b = addmf(b, 'output',1, 'Standard', 'gaussmf', [10 50]);
b = addmf(b, 'output',1, 'Good', 'gaussmf', [10 75]);
b = addmf(b, 'output',1, 'Very good', 'gaussmf', [10 100]);


rule1 = [1 1 1 1 1]; %Very low CC
rule2 = [1 2 2 1 1];
rule3 = [1 3 3 1 1];
rule4 = [1 4 4 1 1];
rule5 = [1 5 5 1 1];

rule6 = [2 1 1 1 1]; %Low CC
rule7 = [2 2 2 1 1];
rule8 = [2 3 3 1 1];
rule9 = [2 4 4 1 1];
rule10 = [2 5 5 1 1];

rule11 = [3 1 2 1 1]; %Average CC
rule12 = [3 2 3 1 1];
rule13 = [3 3 4 1 1];
rule14 = [3 4 4 1 1];
rule15 = [3 5 5 1 1];

rule16 = [4 1 2 1 1]; %High CC
rule17 = [4 2 3 1 1];
rule18 = [4 3 4 1 1];
rule19 = [4 4 5 1 1];
rule20 = [4 5 5 1 1];

rule21 = [5 1 2 1 1]; %Very high CC
rule22 = [5 2 3 1 1];
rule23 = [5 3 4 1 1];
rule24 = [5 4 5 1 1];
rule25 = [5 5 5 1 1];

ruleListB = [rule1; rule2; rule3; rule4;...
    rule5; rule6; rule7; rule8; rule9;...
    rule10; rule11; rule12; rule13;...
    rule14; rule15; rule16; rule17;...
    rule18; rule19; rule20; rule21;...
    rule22; rule23; rule24; rule25;];


b = addrule(b,ruleListB);

fprintf('\nHealer machine:\n\n')

rule = showrule(b)

%defuzzification methods
b.defuzzMethod = 'centroid';

for i=1:size(healerData,1)
        evalplayer = evalfis([healerData(i, 1), healerData(i, 2)], b);
        fprintf('%d) In(1): %.2f, In(2) %.2f,  => Out: %.2f \n',i,healerData(i, 1),healerData(i, 2), evalplayer);  
        xlswrite(filename, evalplayer, 3 , sprintf('C%d',i+1));
        xlswrite(filename, evalplayer, 5 , sprintf('C%d',i+1));
end
%##########################################################################

%********************

%   Damage Machine

%********************
dmgData = xlsread(filename, 4);
c = newfis('Damage');

c = addvar(c, 'input', 'CC Score', [0 75]);

c = addmf(c, 'input',1, 'Very low', 'trimf', [-5 0 7.5]);
c = addmf(c, 'input',1, 'Low', 'gaussmf', [7.5 18.75]);
c = addmf(c, 'input',1, 'Average', 'gaussmf', [7.5 37.5]);
c = addmf(c, 'input',1, 'High', 'gaussmf', [7.5 56.25]);
c = addmf(c, 'input',1, 'Very high', 'trimf', [60 75 80]);


c = addvar(c, 'input', 'Damage', [0 100000]);

c = addmf(c, 'input',2, 'Very low', 'trimf', [-5 0 10000]);
c = addmf(c, 'input',2, 'Low', 'gaussmf', [10000 25000]);
c = addmf(c, 'input',2, 'Average', 'gaussmf', [10000 50000]);
c = addmf(c, 'input',2, 'High', 'gaussmf', [10000 75000]);
c = addmf(c, 'input',2, 'Very high', 'trimf', [80000 100000, 100005]);

c = addvar(c,'output', 'Damage Ability', [0 100]);

c = addmf(c, 'output',1, 'Very poor', 'gaussmf', [10 0]);
c = addmf(c, 'output',1, 'Poor', 'gaussmf', [10 25]);
c = addmf(c, 'output',1, 'Standard', 'gaussmf', [10 50]);
c = addmf(c, 'output',1, 'Good', 'gaussmf', [10 75]);
c = addmf(c, 'output',1, 'Very good', 'gaussmf', [10 100]);


rule1 = [1 1 1 1 1]; %Very low CC
rule2 = [1 2 2 1 1];
rule3 = [1 3 3 1 1];
rule4 = [1 4 4 1 1];
rule5 = [1 5 5 1 1];

rule6 = [2 1 1 1 1]; %Low CC
rule7 = [2 2 2 1 1];
rule8 = [2 3 3 1 1];
rule9 = [2 4 4 1 1];
rule10 = [2 5 5 1 1];

rule11 = [3 1 2 1 1]; %Average CC
rule12 = [3 2 3 1 1];
rule13 = [3 3 4 1 1];
rule14 = [3 4 4 1 1];
rule15 = [3 5 5 1 1];

rule16 = [4 1 2 1 1]; %High CC
rule17 = [4 2 3 1 1];
rule18 = [4 3 4 1 1];
rule19 = [4 4 5 1 1];
rule20 = [4 5 5 1 1];

rule21 = [5 1 2 1 1]; %Very high CC
rule22 = [5 2 3 1 1];
rule23 = [5 3 4 1 1];
rule24 = [5 4 5 1 1];
rule25 = [5 5 5 1 1];

ruleListC = [rule1; rule2; rule3; rule4;
    rule5; rule6; rule7; rule8; rule9;
    rule10; rule11; rule12; rule13;
    rule14; rule15; rule16; rule17;
    rule18; rule19; rule20; rule21;
    rule22; rule23; rule24; rule25;];


c = addrule(c,ruleListC);

fprintf('\nDamage machine:\n\n')

rule = showrule(c)

%defuzzification methods
a.defuzzMethod = 'centroid';

for i=1:size(dmgData,1)
        evalplayer = evalfis([dmgData(i, 1), dmgData(i, 2)], c);
        fprintf('%d) In(1): %.2f, In(2) %.2f,  => Out: %.2f \n',i,dmgData(i, 1),dmgData(i, 2), evalplayer);  
        xlswrite(filename, evalplayer, 4 , sprintf('C%d',i+1));
        xlswrite(filename, evalplayer, 5 , sprintf('D%d',i+1));
end

%##########################################################################

%*************************

%   Helpfulness Machine

%*************************
[helpData,role,~] = xlsread(filename, 5);

fprintf('\nHelpfulness machine:\n\n')

for i=1:size(helpData,1)
    dmg=strcmp(role(i+1,1),'dmg');
    sup=strcmp(role(i+1,1),'sup');
    if sup
        fprintf('\nRow %d belongs to a SUP support\n',i)
        if helpData(i,1)>=helpData(i,2)
            supportStrength = helpData(i,1);
            if supportStrength==helpData(i,2)
                fprintf("Tank Ability and Healer Ability equal: %.2f carried forward\n", supportStrength);
            else
                fprintf("Tank Ability greater than Healer Ability: %.2f carried forward\n", supportStrength);
            end
        elseif helpData(i,1)<helpData(i,2)
            supportStrength = helpData(i,2);
            fprintf("Healer Ability greater than Tank Ability: %.2f carried forward\n", supportStrength);
        end        
        xlswrite('Support_Rank.xls', supportStrength, 5 , sprintf('E%d',i+1));        
        xlswrite('Support_Rank.xls', supportStrength, 6 , sprintf('A%d',i+1));    
    elseif dmg
        supportStrength = helpData(i,3);
        fprintf('\nRow %d belongs to a DPS support\nDamage Ability: %.2f carried forward\n',i, supportStrength)  
        xlswrite('Support_Rank.xls', supportStrength, 5 , sprintf('E%d',i+1));        
        xlswrite('Support_Rank.xls', supportStrength, 6 , sprintf('A%d',i+1));
    end
end

%##########################################################################

%***************************

%   Support Effectiveness Machine

%***************************
d = newfis('Support Effectiveness');
suppData = xlsread(filename, 6);

d = addvar(d, 'input', 'Helpfulness', [0 100]);

d = addmf(d, 'input',1, 'Very poor', 'gaussmf', [10 0]);
d = addmf(d, 'input',1, 'Poor', 'gaussmf', [10 25]);
d = addmf(d, 'input',1, 'Standard', 'gaussmf', [10 50]);
d = addmf(d, 'input',1, 'Good', 'gaussmf', [10 75]);
d = addmf(d, 'input',1, 'Very good', 'gaussmf', [10 100]);

d = addvar(d, 'input', 'Vision Score', [0 110]);

d = addmf(d, 'input',2, 'Very low', 'trimf', [-5 0 11]);
d = addmf(d, 'input',2, 'Low', 'gaussmf', [11 27.5]);
d = addmf(d, 'input',2, 'Average', 'gaussmf', [11 55]);
d = addmf(d, 'input',2, 'High', 'gaussmf', [11 82.5]);
d = addmf(d, 'input',2, 'Very high', 'trimf', [88 110 115]);

d = addvar(d, 'input', 'Assists', [0 50]);

d = addmf(d, 'input',3, 'Very low', 'trimf', [-5 0 5]);
d = addmf(d, 'input',3, 'Low', 'gaussmf', [5 12.5]);
d = addmf(d, 'input',3, 'Average', 'gaussmf', [5 25]);
d = addmf(d, 'input',3, 'High', 'gaussmf', [5 37.5]);
d = addmf(d, 'input',3, 'Very high', 'trimf', [40 50 55]);

d = addvar(d, 'input', 'Deaths', [0 30]);

d = addmf(d, 'input',4, 'Very low', 'trimf', [-5 0 3]);
d = addmf(d, 'input',4, 'Low', 'gaussmf', [3 7.5]);
d = addmf(d, 'input',4, 'Average', 'gaussmf', [3 15]);
d = addmf(d, 'input',4, 'High', 'gaussmf', [3 22.5]);
d = addmf(d, 'input',4, 'Very high', 'trimf', [24 30 35]);

d = addvar(d,'output', 'Support Effectiveness', [0 100]);

d = addmf(d, 'output',1, 'Very poor', 'gaussmf', [10 0]);
d = addmf(d, 'output',1, 'Poor', 'gaussmf', [10 25]);
d = addmf(d, 'output',1, 'Standard', 'gaussmf', [10 50]);
d = addmf(d, 'output',1, 'Good', 'gaussmf', [10 75]);
d = addmf(d, 'output',1, 'Very good', 'gaussmf', [10 100]);

fprintf('\nSupport Effectiveness machine:\n\n')

first = 1;
rCount = 1;
for hlp=1:5
   for vis=1:5
       for ast=1:5
          for dth=1:5
              suppValue = ((hlp+vis+ast)-dth);             
              if (suppValue>2) && (suppValue<= 4)
                    rTemp = [hlp vis ast dth 2 1 1];
                    fprintf('Rule %d: hlp: %d, vis: %d, ast: %d, dth: %d = o:2\n', rCount, hlp, vis, ast, dth);
                    if first == 1
                        ruleListD = rTemp;
                        fprintf('First rule added\n\n')
                    else        
                        ruleListD = [ruleListD; rTemp];
                        fprintf('Rule %d added\n\n', rCount);
                    end 
              elseif (suppValue>4) && (suppValue<= 7)
                    rTemp = [hlp vis ast dth 3 1 1];
                    fprintf('Rule %d: hlp: %d, vis: %d, ast: %d, dth: %d = o:3\n', rCount, hlp, vis, ast, dth);
                    if first == 1
                        ruleListD = rTemp;
                        fprintf('First rule added\n\n')
                    else        
                        ruleListD = [ruleListD; rTemp];
                        fprintf('Rule %d added\n\n', rCount);
                    end 
              elseif (suppValue>7) && (suppValue<= 9)
                    rTemp = [hlp vis ast dth 4 1 1];
                    fprintf('Rule %d: hlp: %d, vis: %d, ast: %d, dth: %d = o:4\n', rCount, hlp, vis, ast, dth);
                    if first == 1
                        ruleListD = rTemp;
                        fprintf('First rule added\n\n')
                    else        
                        ruleListD = [ruleListD; rTemp];
                        fprintf('Rule %d added\n\n', rCount);
                    end 
              elseif suppValue>9
                    rTemp = [hlp vis ast dth 5 1 1];
                    fprintf('Rule %d: hlp: %d, vis: %d, ast: %d, dth: %d = o:5\n', rCount, hlp, vis, ast, dth);
                    if first == 1
                        ruleListD = rTemp;
                        fprintf('First rule added\n\n')
                    else        
                        ruleListD = [ruleListD; rTemp];
                        fprintf('Rule %d added\n\n', rCount);
                    end 
              else
                    rTemp = [hlp vis ast dth 1 1 1];
                    fprintf('Rule %d: hlp: %d, vis: %d, ast: %d, dth: %d = o:1\n', rCount, hlp, vis, ast, dth);
                    if first == 1
                        ruleListD = rTemp;
                        fprintf('First rule added\n\n')
                    else        
                        ruleListD = [ruleListD; rTemp];
                        fprintf('Rule %d added\n\n', rCount);
                    end 
              end
              first = 0;
              rCount = rCount + 1;
          end
       end
   end
end

d = addrule(d,ruleListD)

fprintf('\nSupport Effectiveness machine:\n\n')

rule = showrule(d)

%defuzzification methods
d.defuzzMethod = 'centroid';

for i=1:size(suppData,1)
        evalplayer = evalfis([suppData(i, 1), suppData(i, 2), suppData(i, 3), suppData(i, 4)], d);
        fprintf('%d) In(1): %.2f, In(2) %.2f, In(3) %.2f, In(4) %.2f  => Out: %.2f \n',i,suppData(i, 1),suppData(i, 2),suppData(i,3),suppData(i,4), evalplayer);  
        xlswrite(filename, evalplayer, 6 , sprintf('E%d',i+1));
        xlswrite(filename, evalplayer, 7 , sprintf('A%d',i+1));
end

%##########################################################################

%**************************

%   Support Rank Machine

%**************************
rankData = xlsread(filename, 7);

e = newfis('Support Rank');

e = addvar(e, 'input', 'Support Effectiveness', [0 100]);
e = addmf(e, 'input',1, 'Very poor', 'gaussmf', [10 0]);
e = addmf(e, 'input',1, 'Poor', 'gaussmf', [10 25]);
e = addmf(e, 'input',1, 'Standard', 'gaussmf', [10 50]);
e = addmf(e, 'input',1, 'Good', 'gaussmf', [10 75]);
e = addmf(e, 'input',1, 'Very good', 'gaussmf', [10 100]);


e = addvar(e, 'input', 'Game Duration', [0 80]);

e = addmf(e, 'input',2, 'Very short', 'gaussmf', [6 0]);
e = addmf(e, 'input',2, 'Short', 'gaussmf', [6 15]);
e = addmf(e, 'input',2, 'Average', 'gaussmf', [6 30]);
e = addmf(e, 'input',2, 'Long', 'gaussmf', [6 45]);
e = addmf(e, 'input',2, 'Very long', 'gaussmf', [6 60]);


e = addvar(e,'output', 'Support Rank', [0 100]);

e = addmf(e, 'output',1, 'D', 'gaussmf', [10 0]);
e = addmf(e, 'output',1, 'C', 'gaussmf', [10 25]);
e = addmf(e, 'output',1, 'B', 'gaussmf', [10 50]);
e = addmf(e, 'output',1, 'A', 'gaussmf', [10 75]);
e = addmf(e, 'output',1, 'S', 'gaussmf', [10 100]);


rule1 = [1 1 3 1 1]; %Very low rating
rule2 = [1 2 3 1 1];
rule3 = [1 3 2 1 1];
rule4 = [1 4 1 1 1];
rule5 = [1 5 1 1 1];

rule6 = [2 1 4 1 1]; %Low rating
rule7 = [2 2 3 1 1];
rule8 = [2 3 3 1 1];
rule9 = [2 4 2 1 1];
rule10 = [2 5 1 1 1];

rule11 = [3 1 4 1 1]; %Average rating
rule12 = [3 2 4 1 1];
rule13 = [3 3 3 1 1];
rule14 = [3 4 3 1 1];
rule15 = [3 5 2 1 1];

rule16 = [4 1 5 1 1]; %High rating
rule17 = [4 2 4 1 1];
rule18 = [4 3 4 1 1];
rule19 = [4 4 3 1 1];
rule20 = [4 5 3 1 1];

rule21 = [5 1 5 1 1]; %Very high rating
rule22 = [5 2 5 1 1];
rule23 = [5 3 4 1 1];
rule24 = [5 4 4 1 1];
rule25 = [5 5 3 1 1];

ruleListE = [rule1; rule2; rule3; rule4; rule5;...
    rule6; rule7; rule8; rule9; rule10; rule11;...
    rule12; rule13; rule14; rule15; rule16;...
    rule17; rule18; rule19; rule20; rule21;...
    rule22; rule23; rule24; rule25;];

e = addrule(e,ruleListE);

fprintf('\nSupport Rank machine:\n\n')

rule = showrule(e)

%defuzzification methods
e.defuzzMethod = 'centroid';

for i=1:size(rankData,1)
        evalplayer = evalfis([rankData(i, 1), rankData(i, 2)], e);
        fprintf('%d) In(1): %.2f, In(2) %.2f,  => Out: %.2f \n',i,rankData(i, 1),rankData(i, 2), evalplayer);  
        xlswrite(filename, evalplayer, 7 , sprintf('C%d',i+1));
end
%##########################################################################

%*******************

%   Graph drawing

%*******************

figure(1);
subplot(2,2,1), plotmf(a,'input',1);
subplot(2,2,2), plotmf(a,'input',2);
subplot(2,2,[3,4]), plotmf(a,'output',1);

figure(2);
subplot(2,2,1), plotmf(b,'input',1);
subplot(2,2,2), plotmf(b,'input',2);
subplot(2,2,[3,4]), plotmf(b,'output',1);

figure(3);
subplot(2,2,1), plotmf(c,'input',1);
subplot(2,2,2), plotmf(c,'input',2);
subplot(2,2,[3,4]), plotmf(c,'output',1);

figure(4);
subplot(3,2,1), plotmf(d,'input',1);
subplot(3,2,2), plotmf(d,'input',2);
subplot(3,2,3), plotmf(d,'input',3);
subplot(3,2,4), plotmf(d,'input',4);
subplot(3,2,[5,6]), plotmf(d,'output',1);

figure(5);
subplot(2,2,1), plotmf(e,'input',1);
subplot(2,2,2), plotmf(e,'input',2);
subplot(2,2,[3,4]), plotmf(e,'output',1);