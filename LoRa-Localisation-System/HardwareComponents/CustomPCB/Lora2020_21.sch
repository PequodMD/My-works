EESchema Schematic File Version 4
EELAYER 30 0
EELAYER END
$Descr A4 11693 8268
encoding utf-8
Sheet 1 1
Title ""
Date "lun. 30 mars 2015"
Rev ""
Comp ""
Comment1 ""
Comment2 ""
Comment3 ""
Comment4 ""
$EndDescr
Text Label 8950 1450 1    60   ~ 0
Vin
Text Label 9350 1450 1    60   ~ 0
IOREF
Text Label 8900 2500 0    60   ~ 0
A0
Text Label 8900 2600 0    60   ~ 0
A1
Text Label 8900 2700 0    60   ~ 0
A2
Text Label 8900 2800 0    60   ~ 0
A3
Text Label 8900 2900 0    60   ~ 0
A4(SDA)
Text Label 8900 3000 0    60   ~ 0
A5(SCL)
Text Label 10550 3000 0    60   ~ 0
0(Rx)
Text Label 10550 2800 0    60   ~ 0
2
Text Label 10550 2900 0    60   ~ 0
1(Tx)
Text Label 10550 2700 0    60   ~ 0
3(**)
Text Label 10550 2600 0    60   ~ 0
4
Text Label 10550 2500 0    60   ~ 0
5(**)
Text Label 10550 2300 0    60   ~ 0
7
Text Label 10550 2100 0    60   ~ 0
8
Text Label 10550 1900 0    60   ~ 0
10(**/SS)
Text Label 10550 1800 0    60   ~ 0
11(**/MOSI)
Text Label 10550 1700 0    60   ~ 0
12(MISO)
Text Label 10550 1600 0    60   ~ 0
13(SCK)
Text Label 10550 1400 0    60   ~ 0
AREF
NoConn ~ 9400 1600
Text Label 10550 1300 0    60   ~ 0
A4(SDA)
Text Label 10550 1200 0    60   ~ 0
A5(SCL)
Text Notes 10850 1000 0    60   ~ 0
Holes
Text Notes 8550 750  0    60   ~ 0
Shield for Arduino that uses\nthe same pin disposition\nlike "Uno" board Rev 3.
$Comp
L Connector_Generic:Conn_01x08 P1
U 1 1 56D70129
P 9600 1900
F 0 "P1" H 9600 2350 50  0000 C CNN
F 1 "Power" V 9700 1900 50  0000 C CNN
F 2 "Socket_Arduino_Uno:Socket_Strip_Arduino_1x08" V 9750 1900 20  0000 C CNN
F 3 "" H 9600 1900 50  0000 C CNN
	1    9600 1900
	1    0    0    -1  
$EndComp
Text Label 8650 1800 0    60   ~ 0
Reset
$Comp
L power:+3.3V #PWR01
U 1 1 56D70538
P 9150 1450
F 0 "#PWR01" H 9150 1300 50  0001 C CNN
F 1 "+3.3V" V 9150 1700 50  0000 C CNN
F 2 "" H 9150 1450 50  0000 C CNN
F 3 "" H 9150 1450 50  0000 C CNN
	1    9150 1450
	1    0    0    -1  
$EndComp
$Comp
L power:+5V #PWR02
U 1 1 56D707BB
P 9050 1350
F 0 "#PWR02" H 9050 1200 50  0001 C CNN
F 1 "+5V" V 9050 1550 50  0000 C CNN
F 2 "" H 9050 1350 50  0000 C CNN
F 3 "" H 9050 1350 50  0000 C CNN
	1    9050 1350
	1    0    0    -1  
$EndComp
$Comp
L power:GND #PWR03
U 1 1 56D70CC2
P 9300 3150
F 0 "#PWR03" H 9300 2900 50  0001 C CNN
F 1 "GND" H 9300 3000 50  0000 C CNN
F 2 "" H 9300 3150 50  0000 C CNN
F 3 "" H 9300 3150 50  0000 C CNN
	1    9300 3150
	1    0    0    -1  
$EndComp
$Comp
L power:GND #PWR04
U 1 1 56D70CFF
P 10300 3150
F 0 "#PWR04" H 10300 2900 50  0001 C CNN
F 1 "GND" H 10300 3000 50  0000 C CNN
F 2 "" H 10300 3150 50  0000 C CNN
F 3 "" H 10300 3150 50  0000 C CNN
	1    10300 3150
	1    0    0    -1  
$EndComp
$Comp
L Connector_Generic:Conn_01x06 P2
U 1 1 56D70DD8
P 9600 2700
F 0 "P2" H 9600 2300 50  0000 C CNN
F 1 "Analog" V 9700 2700 50  0000 C CNN
F 2 "Socket_Arduino_Uno:Socket_Strip_Arduino_1x06" V 9750 2750 20  0000 C CNN
F 3 "" H 9600 2700 50  0000 C CNN
	1    9600 2700
	1    0    0    -1  
$EndComp
$Comp
L Connector_Generic:Conn_01x01 P5
U 1 1 56D71177
P 10800 650
F 0 "P5" V 10900 650 50  0000 C CNN
F 1 "CONN_01X01" V 10900 650 50  0001 C CNN
F 2 "Socket_Arduino_Uno:Arduino_1pin" H 10721 724 20  0000 C CNN
F 3 "" H 10800 650 50  0000 C CNN
	1    10800 650 
	0    -1   -1   0   
$EndComp
$Comp
L Connector_Generic:Conn_01x01 P6
U 1 1 56D71274
P 10900 650
F 0 "P6" V 11000 650 50  0000 C CNN
F 1 "CONN_01X01" V 11000 650 50  0001 C CNN
F 2 "Socket_Arduino_Uno:Arduino_1pin" H 10900 650 20  0001 C CNN
F 3 "" H 10900 650 50  0000 C CNN
	1    10900 650 
	0    -1   -1   0   
$EndComp
$Comp
L Connector_Generic:Conn_01x01 P7
U 1 1 56D712A8
P 11000 650
F 0 "P7" V 11100 650 50  0000 C CNN
F 1 "CONN_01X01" V 11100 650 50  0001 C CNN
F 2 "Socket_Arduino_Uno:Arduino_1pin" V 11000 650 20  0001 C CNN
F 3 "" H 11000 650 50  0000 C CNN
	1    11000 650 
	0    -1   -1   0   
$EndComp
$Comp
L Connector_Generic:Conn_01x01 P8
U 1 1 56D712DB
P 11100 650
F 0 "P8" V 11200 650 50  0000 C CNN
F 1 "CONN_01X01" V 11200 650 50  0001 C CNN
F 2 "Socket_Arduino_Uno:Arduino_1pin" H 11024 572 20  0000 C CNN
F 3 "" H 11100 650 50  0000 C CNN
	1    11100 650 
	0    -1   -1   0   
$EndComp
NoConn ~ 10800 850 
NoConn ~ 10900 850 
NoConn ~ 11000 850 
NoConn ~ 11100 850 
$Comp
L Connector_Generic:Conn_01x08 P4
U 1 1 56D7164F
P 10000 2600
F 0 "P4" H 10000 2100 50  0000 C CNN
F 1 "Digital" V 10100 2600 50  0000 C CNN
F 2 "Socket_Arduino_Uno:Socket_Strip_Arduino_1x08" V 10150 2550 20  0000 C CNN
F 3 "" H 10000 2600 50  0000 C CNN
	1    10000 2600
	-1   0    0    -1  
$EndComp
Wire Notes Line
	8525 825  9925 825 
Wire Notes Line
	9925 825  9925 475 
Wire Wire Line
	9350 1450 9350 1700
Wire Wire Line
	9350 1700 9400 1700
Wire Wire Line
	9400 1900 9150 1900
Wire Wire Line
	9400 2000 9050 2000
Wire Wire Line
	9400 2300 8950 2300
Wire Wire Line
	9400 2100 9300 2100
Wire Wire Line
	9400 2200 9300 2200
Connection ~ 9300 2200
Wire Wire Line
	8950 2300 8950 1450
Wire Wire Line
	9050 2000 9050 1350
Wire Wire Line
	9150 1900 9150 1450
Wire Wire Line
	9400 2500 8900 2500
Wire Wire Line
	9400 2600 8900 2600
Wire Wire Line
	9400 2700 8900 2700
Wire Wire Line
	9400 2800 8900 2800
Wire Wire Line
	9400 2900 8900 2900
Wire Wire Line
	9400 3000 8900 3000
$Comp
L Connector_Generic:Conn_01x10 P3
U 1 1 56D721E0
P 10000 1600
F 0 "P3" H 10000 2150 50  0000 C CNN
F 1 "Digital" V 10100 1600 50  0000 C CNN
F 2 "Socket_Arduino_Uno:Socket_Strip_Arduino_1x10" V 10150 1600 20  0000 C CNN
F 3 "" H 10000 1600 50  0000 C CNN
	1    10000 1600
	-1   0    0    -1  
$EndComp
Wire Wire Line
	10200 2100 10550 2100
Wire Wire Line
	10200 2000 10550 2000
Wire Wire Line
	10200 1900 10550 1900
Wire Wire Line
	10200 1800 10550 1800
Wire Wire Line
	10200 1700 10550 1700
Wire Wire Line
	10200 1600 10550 1600
Wire Wire Line
	10200 1400 10550 1400
Wire Wire Line
	10200 1300 10550 1300
Wire Wire Line
	10200 1200 10550 1200
Wire Wire Line
	10200 3000 10550 3000
Wire Wire Line
	10200 2900 10550 2900
Wire Wire Line
	10200 2800 10550 2800
Wire Wire Line
	10200 2700 10550 2700
Wire Wire Line
	10200 2600 10550 2600
Wire Wire Line
	10200 2500 10550 2500
Wire Wire Line
	10200 2400 10550 2400
Wire Wire Line
	10200 2300 10550 2300
Wire Wire Line
	10200 1500 10300 1500
Wire Wire Line
	10300 1500 10300 3150
Wire Wire Line
	9300 2100 9300 2200
Wire Wire Line
	9300 2200 9300 3150
Wire Notes Line
	8500 500  8500 3450
Wire Notes Line
	8500 3450 11200 3450
Wire Wire Line
	9400 1800 8650 1800
Text Notes 9700 1600 0    60   ~ 0
1
Wire Notes Line
	11200 1000 10700 1000
Wire Notes Line
	10700 1000 10700 500 
Wire Wire Line
	6350 3950 6550 3950
$Comp
L power:GND #PWR0101
U 1 1 5F67DB54
P 6550 3950
F 0 "#PWR0101" H 6550 3700 50  0001 C CNN
F 1 "GND" V 6555 3822 50  0000 R CNN
F 2 "" H 6550 3950 50  0001 C CNN
F 3 "" H 6550 3950 50  0001 C CNN
	1    6550 3950
	0    -1   -1   0   
$EndComp
$Comp
L power:+5V #PWR0102
U 1 1 5F68054F
P 3250 3650
F 0 "#PWR0102" H 3250 3500 50  0001 C CNN
F 1 "+5V" V 3265 3778 50  0000 L CNN
F 2 "" H 3250 3650 50  0001 C CNN
F 3 "" H 3250 3650 50  0001 C CNN
	1    3250 3650
	0    -1   -1   0   
$EndComp
Text GLabel 5000 4700 0    50   Input ~ 0
9(RST)
Text GLabel 5000 4800 0    50   Output ~ 0
2(DIO0)
Text GLabel 10550 2000 2    50   Output ~ 0
9(RST)
Text Label 10550 2000 0    60   ~ 0
9(**)
Text Label 10550 2400 0    60   ~ 0
6(**)
Text GLabel 5000 5100 0    50   Output ~ 0
3(DIO3)
Text GLabel 10550 2800 2    50   Input ~ 0
2(DIO0)
Text GLabel 10550 2700 2    50   Input ~ 0
3(DIO3)
Text GLabel 10550 1600 2    50   Output ~ 0
13(SCK)
Connection ~ 6350 2700
Wire Wire Line
	6350 3100 6350 3950
Text GLabel 6950 3100 2    50   Output ~ 0
S5
Text GLabel 6950 3000 2    50   Output ~ 0
S4
Text GLabel 6950 2900 2    50   Output ~ 0
S3
Text GLabel 6950 2800 2    50   Output ~ 0
S2
Text GLabel 6950 2700 2    50   Output ~ 0
S1
Text GLabel 6950 2600 2    50   Output ~ 0
S0
Text GLabel 8900 3000 0    50   Input ~ 0
S5
Text GLabel 8900 2900 0    50   Input ~ 0
S4
Text GLabel 8900 2800 0    50   Input ~ 0
S3
Text GLabel 8900 2700 0    50   Input ~ 0
S2
Text GLabel 8900 2600 0    50   Input ~ 0
S1
Text GLabel 8900 2500 0    50   Input ~ 0
S0
Text GLabel 10550 1700 2    50   Output ~ 0
12(MISO)
Text GLabel 10550 1800 2    50   Output ~ 0
11(MOSI)
Text GLabel 10550 1900 2    50   Output ~ 0
10(NSS)
Text GLabel 10550 2100 2    50   Output ~ 0
LED0
Text GLabel 10550 2300 2    50   Output ~ 0
LED1
$Comp
L Device:LED D0
U 1 1 5F6B527E
P 5500 2700
F 0 "D0" H 5493 2445 50  0000 C CNN
F 1 "LED" H 5493 2536 50  0000 C CNN
F 2 "LED_SMD:LED_1206_3216Metric" H 5500 2700 50  0001 C CNN
F 3 "~" H 5500 2700 50  0001 C CNN
	1    5500 2700
	-1   0    0    1   
$EndComp
$Comp
L Device:LED D1
U 1 1 5F6B9E27
P 5500 3100
F 0 "D1" H 5493 2845 50  0000 C CNN
F 1 "LED" H 5493 2936 50  0000 C CNN
F 2 "LED_SMD:LED_1206_3216Metric" H 5500 3100 50  0001 C CNN
F 3 "~" H 5500 3100 50  0001 C CNN
	1    5500 3100
	-1   0    0    1   
$EndComp
Wire Wire Line
	5650 2700 6350 2700
Wire Wire Line
	5650 3100 6350 3100
Wire Wire Line
	6350 2600 6350 2700
Text GLabel 5050 2700 0    50   Input ~ 0
LED0
Text GLabel 5050 3100 0    50   Input ~ 0
LED1
$Comp
L Device:R R0
U 1 1 5F6D8BEF
P 5200 2700
F 0 "R0" V 4993 2700 50  0000 C CNN
F 1 "R" V 5084 2700 50  0000 C CNN
F 2 "Resistor_SMD:R_1206_3216Metric" V 5130 2700 50  0001 C CNN
F 3 "~" H 5200 2700 50  0001 C CNN
	1    5200 2700
	0    1    1    0   
$EndComp
$Comp
L Device:R R1
U 1 1 5F6D974C
P 5200 3100
F 0 "R1" V 4993 3100 50  0000 C CNN
F 1 "R" V 5084 3100 50  0000 C CNN
F 2 "Resistor_SMD:R_1206_3216Metric" V 5130 3100 50  0001 C CNN
F 3 "~" H 5200 3100 50  0001 C CNN
	1    5200 3100
	0    1    1    0   
$EndComp
Text GLabel 6100 4500 2    50   Input ~ 0
10(NSS)
Text GLabel 6100 4600 2    50   Input ~ 0
11(MOSI)
Text GLabel 6100 4700 2    50   Input ~ 0
12(MISO)
Text GLabel 6100 4800 2    50   Input ~ 0
13(SCK)
Wire Wire Line
	6100 4400 6350 4400
$Comp
L Lora2020_21-rescue:Lora2020_21-cache_Ra-02(sx1278)-Lora2020_21-cache U0
U 1 1 5F702EBE
P 5550 4600
F 0 "U0" H 5550 5025 50  0000 C CNN
F 1 "Lora2020_21-cache_Ra-02(sx1278)" H 5550 4934 50  0000 C CNN
F 2 "Lora2020_21:Ra-02(SX1278)" H 5550 4600 50  0001 C CNN
F 3 "" H 5550 4600 50  0001 C CNN
	1    5550 4600
	1    0    0    -1  
$EndComp
Connection ~ 6350 3100
Wire Wire Line
	6350 2700 6350 2800
$Comp
L Switch:SW_DIP_x06 SW0
U 1 1 5F69BC69
P 6650 2900
F 0 "SW0" H 6650 3467 50  0000 C CNN
F 1 "SW_DIP_x06" H 6650 3376 50  0000 C CNN
F 2 "Button_Switch_THT:SW_DIP_SPSTx06_Slide_9.78x17.42mm_W7.62mm_P2.54mm" H 6650 2900 50  0001 C CNN
F 3 "~" H 6650 2900 50  0001 C CNN
	1    6650 2900
	-1   0    0    -1  
$EndComp
Wire Wire Line
	6350 2800 6350 2900
Connection ~ 6350 2800
Connection ~ 6350 2900
Wire Wire Line
	6350 2900 6350 3000
Connection ~ 6350 3000
Wire Wire Line
	6350 3000 6350 3100
NoConn ~ 10550 2400
NoConn ~ 10550 2500
NoConn ~ 10550 2600
NoConn ~ 10550 3000
NoConn ~ 10550 2900
NoConn ~ 6100 4900
NoConn ~ 6100 5000
NoConn ~ 5000 4900
NoConn ~ 5000 5000
NoConn ~ 6100 5100
NoConn ~ 5000 4400
NoConn ~ 5000 4500
NoConn ~ 10550 1200
NoConn ~ 10550 1300
NoConn ~ 10550 1400
Text Notes 7400 7500 0    50   ~ 0
LoRa shield for Atmega328p
Text Notes 10600 7650 0    50   ~ 0
1
$Comp
L Device:C_Small C2
U 1 1 603C605A
P 3500 3750
F 0 "C2" H 3592 3796 50  0000 L CNN
F 1 "C_Small" H 3592 3705 50  0000 L CNN
F 2 "Capacitor_SMD:C_1206_3216Metric" H 3500 3750 50  0001 C CNN
F 3 "~" H 3500 3750 50  0001 C CNN
	1    3500 3750
	1    0    0    -1  
$EndComp
$Comp
L Device:C_Small C1
U 1 1 603CC66A
P 4650 3850
F 0 "C1" H 4742 3896 50  0000 L CNN
F 1 "C_Small" H 4742 3805 50  0000 L CNN
F 2 "Capacitor_SMD:C_1206_3216Metric" H 4650 3850 50  0001 C CNN
F 3 "~" H 4650 3850 50  0001 C CNN
	1    4650 3850
	1    0    0    -1  
$EndComp
Wire Wire Line
	3250 3650 3500 3650
Wire Wire Line
	3500 3650 3800 3650
Connection ~ 3500 3650
Wire Wire Line
	4400 3650 4500 3650
Wire Wire Line
	4500 3650 4500 4600
Wire Wire Line
	4500 4600 5000 4600
Wire Wire Line
	4500 3650 4650 3650
Wire Wire Line
	4650 3650 4650 3750
Connection ~ 4500 3650
$Comp
L Regulator_Linear:AMS1117-3.3 U1
U 1 1 5FD704F8
P 4100 3650
F 0 "U1" H 4100 3892 50  0000 C CNN
F 1 "AMS1117-3.3" H 4100 3801 50  0000 C CNN
F 2 "Package_TO_SOT_SMD:SOT-223-3_TabPin2" H 4100 3850 50  0001 C CNN
F 3 "http://www.advanced-monolithic.com/pdf/ds1117.pdf" H 4200 3400 50  0001 C CNN
	1    4100 3650
	1    0    0    -1  
$EndComp
Connection ~ 6350 3950
Wire Wire Line
	6350 3950 6350 4400
Wire Wire Line
	3500 3850 3500 3950
Wire Wire Line
	3500 3950 4100 3950
Wire Wire Line
	4100 3950 4650 3950
Connection ~ 4100 3950
Wire Wire Line
	4650 3950 6350 3950
Connection ~ 4650 3950
$EndSCHEMATC
