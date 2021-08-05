# Excel Merging

This utility do simple work - it do summation of multiple same excel files different only in some cell values.
It can process *.xls and *.xlsx files. It can process all sheets.

# For, example:

Have 2 file for summation and one template file for mark cell that should be summed.

![file list](./images/fileList.png)

Inside files:

-file1:

![file1](./images/fileContent1.png)

-file2:

![file2](./images/fileContent2.png)


-and template file:

![file template](./images/fileContentTemplate.png)

The marker for summation cell having value of target cell: **888999** (very rare case).

So just run programm and choose target directory that should be merged (template.xml or template.xlsx must be in that directory!):

![file template](./images/programmExample1.png)

If everifing is ok, you will see view like this:

![file template](./images/programmExample2.png)

Than just click on lower button "start merging" and than you will see that view:

![file template](./images/programmExample3.png)

So, that is all that need for merging. If you click on Directory path label (view image):

![file template](./images/programmExample4.png)

Target directory will be opened via fileExplorer

![file template](./images/programmExample5.png)

The target file now have one more direcotry named "result". Inside that directory pushing merged files.

![file template](./images/programmExample6.png)


All marked cell are summed:

![file template](./images/programmExample7.png)



# Dubugger mode

If need to know more information about sum in resulting cells, you can check checkbox "Add dubug comments in result file?" like in screenshot:

![file template](./images/programmExample10.png)

Pay attention that result file having different type of name:

![file template](./images/programmExample11.png)


Now result file have marking via yellow color background cells. And if your hover top right corner of result cell, your will able to view addition information about result (namefile - source value):

![file template](./images/programmExample12.png)


# P.S.

Please be free to create issue in debug tracker. Or let me know about your results. Thanks!