the programming is mainly done in two part:
1. microcontroller progarmming
2. Java progarmming
___
burn the program to the microcontroller, master for one LoRa device , slave for all other LoRa devices  
connect the master node to the PC that u are using, the java program is not necessary and u can use any serial port communication tool(e.g. Tera Term) to enter command, but u have to do all calculation yourself🌚  
to use the java progarmming, install java on your PC (any version newer or equal to 1.8 will do the job)  
u also need to download the [RXTX library](http://rxtx.qbang.org/wiki/index.php/Download) and put the files(rxtxParallel.dll, rxtxSerial.dll) to your **bin** folder of your java directory  
run cmd, cd to the directory of the `Teminal1_2.jar` file and enter `java -jar Teminal1_2.jar`  
choose the serial port of the master node, than you can enter the command `locateTargetNAlpha 10`, and the `10` is the number of data point will be collected, change it if u want  
after that, enter the x y coordinate of the master node and each anchor nodes in a line  
finally, it will generate a summary xlsx file at **c:\LoRa**, open it to see the result and other data
