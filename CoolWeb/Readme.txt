
CoolWeb - Website Creation Wizard
	author:
	Agam Saran

Table of Contents
-------------------------------------------------------------
	1. Overview
	2. Features
	3. Creating Templates
	4. Creating Pre-defined Pages
	5. License
	6. Support & Credits
-------------------------------------------------------------






PLEASE LEAVE A COMMENT AT PLANET-SOURCE-CODE!
If you like CoolWeb and think it could be useful then please drop a line at Planet-Source-Code. It would encourage me to develop it further and make it even better. A vote would also be nice to rank CoolWeb higher, it will take only a minute or so. I will be very thankful to you for this.




Overview
-------------------------------------------------------------
Do you have ever felt the need of your own professional-looking personal website but do not know even a word of HTML or CSS? Have you ever fell into a situation in which you are required to create a website but do not have much time to do so? If the answer of either question is yes, then CoolWeb is for you.
CoolWeb is a powerful and easy-to-use website creation wizard. CoolWeb will enable you to build professional-looking websites within minutes, you do not even require to be online for it. No need of going through confusing website creation sites and get poor results. With support for templates, CoolWeb is powerful enough to create different modern-looking websites for you. It includes 15 outstanding templates with compact code and more templates can be added quite easily. CoolWeb comes with CoolTemplates a Template Creation Wizard which enables you to create templates very easily. The included templates are designed by respective users of OSWD (www.oswd.org). Please take some time to visit OSWD and the websites of authors of the templates. Aside from templates, CoolWeb also focuses on time taken to build a site. It will take about only 10 minutes to create your ready-to-publish website. By using the term "ready-to-publish", I mean that if you designed your website the way you want in CoolWeb, you do not need to make any changes in order to publish it on the web. Just build and publish! CoolWeb also supports "Projects" so that you can continue from where you last left off. It supports Unicode characters too!
With CoolWeb, creating websites is truly visual. Select templates visually, add and delete pages visually and format your text visually. If you are in a little hurry, and do not have time to write text for your website, CoolWeb also includes some predefined-pages for you. It couldn't have been easier than that! More predefined-pages can be added easily, just drop them in the "pages" folder.



Features
-------------------------------------------------------------
1. 15 exciting and modern templates taken from OSWD (http://www.oswd.org/).
2. Support for more templates.
3. Comes with CoolTemplates, easily create templates.
3. With support for Projects, you can easily continue your work.
4. Support for pre-defined pages helps you to write text for your website. Some pages are included already, more can be added easily.
5. Support for a Multi-lingual GUI.
6. No changes required in order to publish after building in CoolWeb.
7. Everything can be done visually, you do not need to "learn" anything.



Creating Templates
-------------------------------------------------------------
You can use CoolTemplates to create templates for CoolWeb, it is recommended. But if you want to create the templates manually, below is the procedure for doing it. If you can't get it, the included templates can work as examples.

1. Firstly, make a folder with the name of the template in the "templates" folder.

2. Design your template, place all files in the folder you created in Step 1.

3. Create a file called "menuitem.txt" in this folder and place the HTML code for a single menu-item of your template in this file. Now replace the value of "href" attribute with "<!--ItemPath-->" and the name of menu-item with "<!--ItemName-->". For example: if HTML code for a menu-item is:
	<li><a href="index.htm">Home</a></li>
You have to turn it to:
	<li><a href="<!--ItemPath-->"><!--ItemName--></a></li>

4. Create another file called "main.txt" and place all the HTML code of your template in this file. Place the following comments at their places:
	<!--Title-->		Place where you want the title of website, such as between <title> and </title>.
	<!--Author--> 		Place where you want the name of author of website, such as at the footer or in the "author" meta-data.
	<!--PageName-->		Place where you want the name of page, such as at the header.
	<!--PageContent-->		Put where you want the actual content of page.
	<!--MenuItems-->		Replace the HTML code of all the menu-items with this. CoolWeb will refer to "menuitem.txt" for the code of menu-items.
	<!--Year-->		Place where you want the year of creation of website.
	<!--Date-->		Leave where you want the date of creation of website.

Note: These comments are case-sensitive.

5. Create a folder called "files" in the folder you created in Step 1, place all the files that are used by your template, including images and stylesheets, in this folder. Now make necessary changes in "main.txt" for this so that all the files are taken from "files" folder. For Example: if the HTML code of your template contains the following line:
	<img src="mypic.jpg">
You have to turn it to:
	<img src="files/mypic.jpg">

6. Create yet another file called "data.ini" in the folder you created in Step 1, CoolWeb will not show your template in the list if you do not create this file. Now put the following content in this file:
[Data]
Name=(Name of Template)
Author=(Name of Author of Template)
Description=(Short description of Template)
URL=(Website of Author)

7. Now capture screenshot of your template. Save it in the JPG format with two different sizes, 190x190 and 80x80, with the names of "screen_full.jpg" and "screen_thumb.jpg" respectively, in the folder you created in Step 1.

8. Congrats! You have created your template. CoolWeb will read it automatically and it should work fine if you did it right.



Creating Pre-defined Pages
-------------------------------------------------------------
Adding Pre-defined Pages is very easy, just create them and drop them in the "pages" folder. CoolWeb will read and show them automatically.



License
-------------------------------------------------------------
CoolWeb is 100% yours now, you can do whatever pleases you with it. But I am not responsible for anything, absolutely anything, that goes wrong. Use at your own risk!



Support and Credits
-------------------------------------------------------------
CoolWeb has been created and programmed by Agam Saran (http://www.agamsaran.co.nr/). The templates used in CoolWeb are designed by users of OSWD (http://www.oswd.org/). I thank them very very much for creating those outstanding templates. Their names, along with their websites, are written just below respective screenshots in the "Select Template" step. Please take some time to visit their websites and OSWD. The cool-looking toolbar icon set used in CoolWeb (called "Silk Icon Set") has been copyrighted and designed by Mark James (http://www.famfamfam.com/lab/icons/silk/).

The following people have been keen-enough to take part in the coding and development of CoolWeb:

Xpert
For notifying various bugs and providing his excellent "CreateFolder" function as a replacement for "SHCreateDirectoryEx" API.

Pietro Cecchi
For pointing that Unicode support should be added and providing the code (TakeCareOfUnicode function) to accomplish this task.

Aside from them, I would like to thank everybody who voted for me. I really mean it! If you have a question, a comment or a suggestion, feel free to mail me at "contactme@axigenmail.com". I will feel glad to hear your comments or to help you.