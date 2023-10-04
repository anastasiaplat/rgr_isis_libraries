# rgr_isis_libraries
Project for uni in subject "Information system tools" 4 course.

The aim of the project is to learn object models and use libraries of Excel, Word, PowerPoint, Access.

The file Students.cs is a class that represents a set of student's information: name, 4 grades and additional field (according individual number given by prepod) - social work. 
Also it has a method for calculating amount of grants and this is the most useful thing in this project. I think so. 

The file Form1.cs has description for 8 buttons (8 tasks). Let me explain what are they for.

1. This part of code opens Excel file and take data from it about students, then create new Excel file and put there calculated grants amounts for each student.

2. Here the data from that file is becoming a base for creating 2 types of diagram in new Excel file .

3. If you click on this button, it will create Word file with list of students' names and grants.

4. Almost the same as previous, but new data will be as table.

5. This button creates PowerPoint presentation with 4 slides according for that Excel file, of course. The first slide contains title and subtitle. The second one - title with the picture, the another one - a list of names and grants and the last one - a table with the same data.

6. There is another input file: a Word file with random paragraphs. The code creates a presentation with slides as much as count of paragraphs in that file.

7. Now the code will fill the cells of Grants field in Access file with names and grades of students.

8. Finally, the code creates a Word file with list of students' grants according a data from Access file.


That's all.
