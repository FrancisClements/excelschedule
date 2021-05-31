# ExcelSchedule
ExcelSchedule is a Python application that converts a table-written schedule into a graphical schedule using Excel.

![Menu](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/menu.PNG)
![Options](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/options.PNG)

# How to Use the App
## Main Menu
1. Browse your Excel file that contains your schedule. This will be your **Input File**.
> Make sure that your file contains a set of headers. And among those headers, it should at least have the following categories:
> * Time
> * Subject Name or Course Name

> These are examples of valid schedule files.
![Sample1](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/inputfile.PNG)
![Sample2](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/inputfile2.PNG)
2. Insert the name of the **Output Excel file**. This file will be saved at the same location as the input file.

## Customization
1. **Enable Hour List (optional)** inserts the hour number before the time itself. It will add "1" at the left side of your schedule if your start time is at "1:05PM" or "1:45PM". This is useful for some subjects that take place at the same hour.
2. **Enable Header (optional)** lets you add the title/header of the schedule.
3. **Include your Name (optional)** is self-explanatory. Include your name to your schedule.
4. **Include description of the subject (optional)** adds a short description for your subjects. This is usually useful for college-level cases such as:
> * Room Navigation. When subjects have different rooms to enter on that certain day. e.g. On Math Subject, you go to Room 203 on Mondays, while Room 305 on Fridays
> * Irregular Sections. When subjects have different class sections. e.g. On Math you are at the "A-202" class, while on English, you are at the "BA-201" class
5. **Add Day Format (optional)** lets you place which day of the week is the subject at. 
> * On COLUMN menu, you need to specify which column corresponds to the days of the subject.
> * On FORMAT menu, you determine the formatting of the days. 
>   * Initial shows M, T, W, ...
>   * Partial shows Mon, Tues, Wed, ...
>   * Full shows Monday, Tuesday, Wednesday, ...
6. **Time Format** configures when the subject takes place.
> * The first menu contains the formatting of the time for your schedule. It has 3 formats:
>   * 12 hr + AM/PM (1:01PM)
>   * 12 hr + a/p (1:01p)
>   * 24 hr (13:01p)
> * **I need 2 columns to set it** means that in your schedule table, the **time in** and **time out** are seperated. Unchecking it means that the time range are contained in a single cell.
>   * Either way, you need to specify which column corresponds to the subject time.
7. **Color Selection** lets you set the coloring of the subjects to your schedule.
> * Category is what will show up to your schedule. Recommended to be brief and short. e.g. Your subject, Course Code.
> * Font Color menu lets you customize the text of all the category font color.
8. Press **Create Schedule** button to finally create the schedule. It should show that the Excel file has been created.

# Example Output
A little reminder that since it's Excel, you can change the cell size, font style and font size to your liking as a final touch.

**All images are real, used schedules*

![Sched1](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/sched_2ndyr.png)
![Sched2](https://github.com/FrancisClements/excelschedule/blob/master/screenshots/sched_2ndyr2.png)
