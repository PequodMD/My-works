the [arduino IDE](https://www.arduino.cc/en/software) is used for the microcontroller programming  
a [LoRa radio chip library](https://github.com/sandeepmistry/arduino-LoRa) is used in the program, download it and load it to your arduino IDE  
u also need to modify the some private functions to public like `LoRaClass::readRegister(uint8_t address)`  
