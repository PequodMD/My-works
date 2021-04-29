
the PCB is used to provide circuit connection, debuging and addressing function  
design of PCB is done by [KiCAD](https://kicad.org/)  
[**ALL** project files for KiCAD](Lora2020_21.zip)  
[schematic file](Lora2020_21.sch)  
![sch](https://github.com/PequodMD/Upload/blob/278032360a0b57289a7eeacce4fd8d3fa8e4177a/images/imagesForLoRaLocalisationSystem/sch.png)  
[circuit layout file](Lora2020_21.kicad_pcb)  
![layout](https://github.com/PequodMD/Upload/blob/278032360a0b57289a7eeacce4fd8d3fa8e4177a/images/imagesForLoRaLocalisationSystem/layout.png)  
[gerber and drill files](Lora2020_21-gerbers3.zip)  
![gerber](https://github.com/PequodMD/Upload/blob/278032360a0b57289a7eeacce4fd8d3fa8e4177a/images/imagesForLoRaLocalisationSystem/gerber.png)  
[Video for learning how to use KiCAD](https://www.youtube.com/watch?v=vaCVh2SAZY4)  
appearance of the board  
![boardup](https://github.com/PequodMD/Upload/blob/89a961c7ec6871831e518f6c5afcf2171c23f78a/images/imagesForLoRaLocalisationSystem/PCBup.png)
![boarddown](https://github.com/PequodMD/Upload/blob/89a961c7ec6871831e518f6c5afcf2171c23f78a/images/imagesForLoRaLocalisationSystem/PCBdown.jpg)
___
components on top:  
 + 2 LEDs and resistors(for debuging)
 + 6 ways DIP switch(to config address of the device)
 + a 5V to 3.3V regulator and two capacitors(power supply for Ra-02 module)
___
to fabricate the PCB, submit the GERBER files(Front and Bottom of copper, mask and silks layers) with the drill files(zip in one) to the following website  
[OSHPARK(for final PCB version)](https://oshpark.com/)  
[JLPCB(for prototype PCB)](https://jlcpcb.com/)
