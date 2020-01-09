

%input data from 'Grades' sheet
opts = detectImportOptions('GPA.xlsx');
opts.Sheet = 'Grades';
opts.SelectedVariableNames = [1:8]; 
opts.DataRange = '2:9';
grades = readtable('GPA.xlsx',opts); 
disp(grades); 

%input data from 'Credits' sheet
opts = detectImportOptions('GPA.xlsx');
opts.Sheet = 'Credits';
opts.SelectedVariableNames = [1:7]; 
opts.DataRange = '2:2';
credits = readtable('GPA.xlsx',opts);
disp(credits); 

totalCredits = sum(credits{1:end,2:end},2)
creditsVector = table2array(credits(1,2:end))


for index = 1:8 
    gradesVector = table2array(grades(index,2:7)); 
    dotProduct = dot(creditsVector, gradesVector); 
    studentGPA = dotProduct / totalCredits;
    disp(studentGPA);
    grades{index,8} = {studentGPA}; 
    %disp(grades); 

end
disp(grades); 

writetable(grades,'GPA.xlsx')

%input data from 'Grades' sheet to verify
opts = detectImportOptions('GPA.xlsx');
opts.Sheet = 'Grades';
opts.SelectedVariableNames = [1:8]; 
opts.DataRange = '1:9';
finalgrades = readtable('GPA.xlsx',opts); 
disp(finalgrades); 







 
    
    
