#include<LoRa.h>

const uint8_t ping = 1;
const uint8_t setNodeType = 2;
const uint8_t setNodeBW = 3;
const uint8_t setNodeCR = 4;
const uint8_t setNodeSF = 5;
const uint8_t setNodeTxPower = 6;
const uint8_t setNodeLNA = 7;
const uint8_t timeStampAdjustment = 8;
const uint8_t discoverNodes = 9;
const uint8_t setNodeOrder = 10;
const uint8_t startTargetToBaseTransmition = 11;
const uint8_t baseToMasterTransmition = 12;
const uint8_t targetToBaseTransmition = 13;
const uint8_t setNodePL = 14;
const uint8_t setNodeSW = 15;
const uint8_t nodeOnOff = 16;

const boolean targetNode = true;
const boolean baseStationNode = false;
const boolean node2Node = true;

const boolean master2Node = false;
const uint8_t BW250E3 = 0;
const uint8_t BW125E3 = 1;
const uint8_t BW62P5E3 = 2;
const uint8_t BW41P7E3 = 3;
const uint8_t BW31P25E3 = 4;
const uint8_t BW20P8E3 = 5;
const uint8_t BW15P6E3 = 6;
const uint8_t BW10P4E3 = 7;
const uint8_t BW7P8E3 = 8;
const uint8_t LNAG0 = 0;
const uint8_t LNAG1MAX = 1;
const uint8_t LNAG2 = 2;
const uint8_t LNAG3 = 3;
const uint8_t LNAG4 = 4;
const uint8_t LNAG5 = 5;
const uint8_t LNAG6MIN = 6;
const uint8_t LNAAuto = 8;
const boolean nodeOn = true;
const boolean nodeOff = false;
const uint8_t messagelength = 13;
const long minTransmitTime = 36;
const double master2NodeBW = 41.7E3;
const uint8_t master2NodeSpreadingFactor = 9;
const uint8_t master2NodeCodingRate = 8;
const unsigned short master2NodePreambleLength = 0x0F;
const uint8_t master2NodeSyncWord = 0xFF;
const uint8_t master2NodeTxPower = 17;
const uint8_t master2NodeLNAStatus = LNAAuto;

uint8_t thisNodeAddr;
const uint8_t masterNodeAddr = 0x00;
const uint8_t boardCastAddr = 0xFF;

unsigned short miniSecond;
unsigned long second;
boolean nodeType;
boolean nodeMode = master2Node;
boolean OnOff = nodeOn;
uint8_t nodeOrder;
unsigned long timeStamp;
uint8_t command = 0;
uint8_t commandPara[10];
double node2NodeBW = 41.7E3;
uint8_t node2NodeSpreadingFactor = 9;
uint8_t node2NodeCodingRate = 5;
unsigned short node2NodePreambleLength = 0x0F;
uint8_t node2NodeSyncWord = 0xAA;
uint8_t node2NodeTxPower = 17;
uint8_t node2NodeLNAStatus = LNAAuto;
long n2nTransmitTime;
long m2nTransmitTime;

//#####################################################################################################
//########################################################setup()######################################
//#####################################################################################################
void setup() 
{
  Serial.begin(9600);
  while(!Serial);
  Serial.println("Slave side");
  int setUp;
  while(true)
  {
    setUp = LoRa.begin(433E6);
    if(setUp)
    {
      Serial.println("Connected to LoRa chips");
      LoraSlaveMasterInitialization();
      timerInitialization();
      sei();
      Serial.println("Address: "+String(thisNodeAddr)+" Initialization completed");
      break;
    }else{
      Serial.println("Failed to connect, retry in 10s");
      delay(10E3);
    }
  }
}
//#####################################################################################################
//########################################################loop()#######################################
//#####################################################################################################
void loop() 
{
  if(nodeMode==node2Node)
  {
    if((timeStamp+500+n2nTransmitTime)<(second*1000+miniSecond))
    {
      nodeMode = master2Node;
      switchToMaster2Node();
      Serial.println("timeout for target Transmition");
      green();
    }
  }
  switch(command)
  {
    case(ping):
    {
      pingResponse();
      command = 0;
      green();
      break;
    }
    case(setNodeType):
    {
      setNodeTypeTo();
      command = 0;
      green();
      break;
    }
    case(setNodeBW):
    {
      setNode2NodeBWTo();
      command = 0;
      green();
      break;
    }
    case(setNodeCR):
    {
      setNode2NodeCRTo();
      command = 0;
      green();
      break;
    }
    case(setNodeSF):
    {
      setNode2NodeSFTo();
      command = 0;
      green();
      break;
    }
    case(setNodeTxPower):
    {
      setNode2NodeTxPowerTo();
      command = 0;
      green();
      break;
    }
    case(setNodeLNA):
    {
      setNode2NodeLNATo();
      command = 0;
      green();
      break;
    }
    case(timeStampAdjustment):
    {
      AdjustmentTimeStamp();
      command = 0;
      green();
      break;
    }
    case(setNodeOrder):
    {
      setNodeOrderTo();
      command = 0;
      green();
      break;
    }
    case(startTargetToBaseTransmition):
    {
      transmitionStarted();
      command = 0;
      break;
    }
    case(targetToBaseTransmition):
    {
      //Serial.println("Start f TX");
      both();
      baseToMaster();
      green();
      command = 0;
      break;
    }
    case(setNodePL):
    {
      setNodePLTo();
      command = 0;
      green();
      break;
    }
    case(setNodeSW):
    {
      setNodeSWTo();
      command = 0;
      green();
      break;
    }
    case(nodeOnOff):
    {
      setNodeOnOff();
      command = 0;
      green();
      break;
    }
    default:
    {
      break;
    }
  }
}
//#####################################################################################################
//########################################################function#####################################
//#####################################################################################################
//start transmition/////////////////////////////////////////////////////////////////////////////////////
void transmitionStarted()
{
  LoRa.idle();
  sei();
  if(!OnOff)
  {
    LoRa.receive();
    return;
  }
  switchToNode2Node();
  if(nodeType==targetNode)
  {
    nodeMode=node2Node;
    targetToBaseTx();
  }else{
    nodeMode=node2Node;
    timeStamp = second*1000+miniSecond;
  }
  LoRa.receive();
}
//target To Base Transmition////////////////////////////////////////////////////////////////////////////
void targetToBaseTx()
{
  LoRa.idle();
  //Serial.println("Start Target TX");
  sei();
  String messageToSend = "";
  messageToSend+=(char)boardCastAddr;
  messageToSend+=(char)thisNodeAddr;
  messageToSend+=(char)targetToBaseTransmition;
  messageToSend+=(char)nodeOrder;
  while(messageToSend.length()<messagelength)
  {
    messageToSend+='~';
  }
  LoRa.beginPacket();
  LoRa.print(messageToSend);
  LoRa.endPacket();
  switchToMaster2Node();
  nodeMode = master2Node;
  LoRa.receive();
}
//Base transmition To Master/////////////////////////////////////////////////////////////////////////////////////
void baseToMaster()
{
  LoRa.idle();
  //Serial.println("Start Base TX");
  sei();
  switchToMaster2Node();
  switchToNode2MasterTx();
  timeStamp = second*1000+miniSecond;
  if(commandPara[0]>nodeOrder)
  {
    //while((second*1000+miniSecond)<(85*nodeOrder+timeStamp));
    delay((m2nTransmitTime+50)*nodeOrder+1);
  }else{
    delay((m2nTransmitTime+50)*(nodeOrder));
  }
  String messageToSend = "";
  messageToSend+=(char)masterNodeAddr;
  messageToSend+=(char)thisNodeAddr;
  messageToSend+=(char)baseToMasterTransmition;
  messageToSend+=(char)commandPara[1];
  messageToSend+=(char)commandPara[2];
  messageToSend+=(char)commandPara[3];
  messageToSend+=(char)commandPara[4];
  messageToSend+=(char)commandPara[5];
  messageToSend+=(char)commandPara[6];
  messageToSend+=(char)commandPara[7];
  while(messageToSend.length()<messagelength)
  {
    messageToSend+='~';
  }
  LoRa.beginPacket();
  LoRa.print(messageToSend);
  LoRa.endPacket();
  switchBackToMaster2Node();
  nodeMode = master2Node;
  LoRa.receive();
}
//set Node to Node Bandwidth/////////////////////////////////////////////////////////////////////////////////////
void setNode2NodeBWTo()
{
  LoRa.idle();
  sei();
  switch(commandPara[0])
  {
    case(BW250E3):
    {
     node2NodeBW = 250E3;
      break;
    }
    case(BW125E3):
    {
      node2NodeBW = 125E3;
      break;
    }
    case(BW62P5E3):
    {
      node2NodeBW = 62.5E3;
      break;
    }
    case(BW41P7E3):
    {
      node2NodeBW = 41.7E3;
      break;
    }case(BW31P25E3):
    {
      node2NodeBW = 31.25E3;
      break;
    }
    case(BW20P8E3):
    {
      node2NodeBW = 20.8E3;
      break;
    }
    case(BW15P6E3):
    {
      node2NodeBW = 15.6E3;
      break;
    }
    case(BW10P4E3):
    {
      node2NodeBW = 10.4E3;
      break;
    }
    case(BW7P8E3):
    {
      node2NodeBW = 7.8E3;
      break;
    }
  }
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  Serial.println("Node BW set to "+String(node2NodeBW));
  LoRa.receive();
}
//set Node to Node Sprending Factor/////////////////////////////////////////////////////////////////////////////////////
void setNode2NodeSFTo()
{
  LoRa.idle();
  sei();
  node2NodeSpreadingFactor = (uint8_t)commandPara[0];
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  Serial.println("Node Spreading Factor set to "+String(node2NodeSpreadingFactor));
  LoRa.receive();
}
//set Node to Node coding rate/////////////////////////////////////////////////////////////////////////////////////
void setNode2NodeCRTo()
{
  LoRa.idle();
  sei();
  node2NodeCodingRate = (uint8_t)commandPara[0];
  Serial.println("Node Coding Rate set to "+String(node2NodeCodingRate));
  LoRa.receive();
}
//set Node to Node Preamble Length/////////////////////////////////////////////////////////////////////////////////////
void setNodePLTo()
{
  LoRa.idle();
  sei();
  node2NodePreambleLength = (unsigned short)commandPara[0]<<8|(uint8_t)commandPara[1];
  Serial.println("Node Preamble Length set to "+String(node2NodePreambleLength));
  LoRa.receive();
}
//set Node to Node sync Word/////////////////////////////////////////////////////////////////////////////////////
void setNodeSWTo()
{
  LoRa.idle();
  sei();
  node2NodeSyncWord = (uint8_t)commandPara[0];
  Serial.println("Node Sync Word set to "+String(node2NodeSyncWord, HEX));
  LoRa.receive();
}
//set Node to Node Tx power/////////////////////////////////////////////////////////////////////////////////////
void setNode2NodeTxPowerTo()
{
  LoRa.idle();
  sei();
  node2NodeTxPower = (uint8_t)commandPara[0];
  Serial.println("Node Transmition Power set to "+String(node2NodeTxPower));
  LoRa.receive();
}
//set Node LNA/////////////////////////////////////////////////////////////////////////////////////
void setNode2NodeLNATo()
{
  LoRa.idle();
  sei();
  node2NodeLNAStatus = (uint8_t)commandPara[0];
  if(0x07&node2NodeLNAStatus==0)
  {
    Serial.println("Node LNA Status set to off");
  }else if(0x08&node2NodeLNAStatus){
    Serial.println("Node LNA Status set to Auto");
  }else{
    Serial.println("Node LNA Status set to G"+String(0x07&node2NodeLNAStatus));
  }
  LoRa.receive();
}
//set Node On or Off//////////////////////////////////////////////////////////////////////////////////
void setNodeOnOff()
{
  LoRa.idle();
  sei();
  OnOff = (uint8_t)commandPara[0];
  if(OnOff)
  {
    Serial.println("Set Node to On");
  }else{
    Serial.println("Set Node to Off");
  }
  LoRa.receive();
}
//set Node Order/////////////////////////////////////////////////////////////////////////////////////
void setNodeOrderTo()
{
  LoRa.idle();
  sei();
  nodeOrder = (uint8_t)commandPara[0];
  Serial.println("Set Local Node Order to : "+String(nodeOrder));
  LoRa.receive();
}
//set Node TimeStamp/////////////////////////////////////////////////////////////////////////////////////
void AdjustmentTimeStamp()
{
  LoRa.idle();
  unsigned long masterTimeStamp = (unsigned long)(commandPara[6]<<24)|(unsigned long)((unsigned short)commandPara[7]<<8)<<8|(unsigned short)commandPara[8]<<8|(uint8_t)commandPara[9];
  unsigned long newCalLocalTimeStamp = masterTimeStamp+52;
  timerInitialization();
  unsigned long localTimeStamp = second*1000+miniSecond;
  miniSecond = (unsigned short)(newCalLocalTimeStamp%1000);
  second = (unsigned long)(newCalLocalTimeStamp/1000);
  sei();
  unsigned long newLocalTimeStamp = second*1000+miniSecond;
  Serial.println("new Local time Stamp: "+String(newLocalTimeStamp));
  Serial.println("Local time Stamp: "+String(localTimeStamp));
  Serial.println("time Stamp different: "+String((signed long)localTimeStamp-(signed long)masterTimeStamp));
  Serial.println("Master time Stamp: "+String(masterTimeStamp));
  Serial.println("new cal Local time Stamp: "+String(newCalLocalTimeStamp));
  LoRa.receive();
}
//set Node type/////////////////////////////////////////////////////////////////////////////////////
void setNodeTypeTo()
{
  LoRa.idle();
  sei();
  if(commandPara[0])
  {
    Serial.println("set to Target");
    nodeType = targetNode;
  }else{
    Serial.println("set to BaseStation");
    nodeType = baseStationNode;
  }
  LoRa.receive();
}
//ping Response/////////////////////////////////////////////////////////////////////////////////////
void pingResponse()
{
  LoRa.idle();
  sei();
  switchToNode2MasterTx();
  String messageToSend = "";
  messageToSend+=(char)masterNodeAddr;
  messageToSend+=(char)thisNodeAddr;
  messageToSend+=(char)ping;
  messageToSend+=(char)commandPara[0];
  messageToSend+=(char)commandPara[1];
  while(messageToSend.length()<messagelength)
  {
    messageToSend+='~';
  }
  LoRa.beginPacket();
  LoRa.print(messageToSend);
  LoRa.endPacket();
  switchBackToMaster2Node();
  LoRa.receive();
}
//switch To Node to Node mode/////////////////////////////////////////////////////////////////////////////////////
void switchToNode2Node()
{
  sei();
  LoRa.setSignalBandwidth(node2NodeBW);
  LoRa.setSpreadingFactor(node2NodeSpreadingFactor);
  LoRa.setCodingRate4(node2NodeCodingRate);
  LoRa.setPreambleLength(node2NodePreambleLength);
  LoRa.setSyncWord(node2NodeSyncWord); 
  LoRa.setTxPower(node2NodeTxPower);
  LoraLNASetting(node2NodeLNAStatus);
  LoRa.disableInvertIQ();
}
//switch To Master to Node mode/////////////////////////////////////////////////////////////////////////////////////
void switchToMaster2Node()
{
  sei();
  LoRa.setSignalBandwidth(master2NodeBW);
  LoRa.setSpreadingFactor(master2NodeSpreadingFactor);
  LoRa.setCodingRate4(master2NodeCodingRate);
  LoRa.setPreambleLength(master2NodePreambleLength);
  LoRa.setSyncWord(master2NodeSyncWord);
  LoRa.setTxPower(master2NodeTxPower);
  LoraLNASetting(master2NodeLNAStatus);
  LoRa.disableInvertIQ();
}
//switch Node to master Tx/////////////////////////////////////////////////////////////////////////////////////
void switchToNode2MasterTx()
{
  sei();
  LoRa.enableInvertIQ();
}
//switch master to node Tx/////////////////////////////////////////////////////////////////////////////////////
void switchBackToMaster2Node()
{
  sei();
  LoRa.disableInvertIQ();
}
void LoraLNASetting(uint8_t LNASetting)
{
  LoRa.writeRegister(0x0C, (0x07&LNASetting)<<5);
  LoRa.writeRegister(0x26, (0x08&LNASetting)>>1|(0xFB&LoRa.readRegister(0x26)));
}
//power//////////////////////////////////////////////////////////////////////
int power(int index)
{
  int result = 1;
  switch(index)
  {
    case(1):
    {
      result = 2;
      break;
    }
    case(2):
    {
      result = 4;
      break;
    }
    case(3):
    {
      result = 8;
      break;
    }
    case(4):
    {
      result = 16;
      break;
    }
    case(5):
    {
      result = 32;
      break;
    }
    default:
    {
      break;
    }
  }
  return result;
}
//#####################################################################################################
//########################################################Interrupt####################################
//#####################################################################################################
//Dio0 RX Interrupt//////////////////////////////////////////////////////////////////////////////////////
void receiveIQR(int size) 
{
  sei();
  String message = "";
  while(LoRa.available())
  {
    char chara = (char)LoRa.read();
    message += chara;
  }
  //Serial.println(message);
  if((uint8_t)message.charAt(0)==thisNodeAddr||(uint8_t)message.charAt(0)==boardCastAddr)
  {
    red();
    //Serial.println((uint8_t)message.charAt(0));
    if((uint8_t)message.charAt(1)==masterNodeAddr)
    {
      command = (uint8_t)message.charAt(2);
      //Serial.println("com: "+String(command));
      if(command==ping)
      {
        short pingRssi= (short)LoRa.packetRssi();
        commandPara[0] = 0xFF&(pingRssi>>8);
        commandPara[1] = 0xFF&pingRssi;
      }else if(command==startTargetToBaseTransmition){
        commandPara[7] = LoRa.readRegister(0x1A);//alpha value
      }else{
        for(int mesageIndex = 3;mesageIndex<13;mesageIndex++)
        {
          commandPara[mesageIndex-3] = (uint8_t)message.charAt(mesageIndex);
        }
      }
    }else{
      command = (uint8_t)message.charAt(2);
      //Serial.println(String(command));
      if(command==targetToBaseTransmition)
      {
        Serial.println("signal cut-in");
        commandPara[0] = (uint8_t)message.charAt(3);//target order number
        commandPara[1] = LoRa.readRegister(0x1A);
        commandPara[2] = LoRa.readRegister(0x19);
        long fer = LoRa.packetFrequencyError();
        commandPara[3] = 0xFF&(fer>>24);
        commandPara[4] = 0xFF&(fer>>16);
        commandPara[5] = 0xFF&(fer>>8);
        commandPara[6] = 0xFF&(fer);
      }
    }
  }
}
//timer Interrupt//////////////////////////////////////////////////////////////////////////////////////
ISR(TIMER1_COMPA_vect)
{
  miniSecond++;
  if(miniSecond==1000)
  {
    miniSecond = 0;
    second++;
  }
}
//#####################################################################################################
//LED
void green()
{
  PORTD|=0x80;//LED1
  PORTB&=0xFE;//LED2 //green only
}
void red()
{
  PORTD&=0x7F;//LED1
  PORTB|=0x01;//LED2 //red only
}
void both()
{
  PORTD|=0x80;//LED1
  PORTB|=0x01;//LED2 //both
}
//########################################################Initialization###############################
//#####################################################################################################
//timer Initialization/////////////////////////////////////////////////////////////////////////////////
void timerInitialization()
{
  TCCR1A = 0;
  TCCR1B = 0x09;
  OCR1A = 0x3E7F;
  TCNT1  = 0;
  TIMSK1 |= (1 << OCIE1A);
}
//LoRa Initialization/////////////////////////////////////////////////////////////////////////////////
void LoraSlaveMasterInitialization()
{
  DDRD&=0x0F;//address Switch
  PORTD|=0xF0;//pullup
  DDRC|=0x07;//LED
  delay(1);

  PORTC|=0x05;//both
  
  PORTC&=0xFB;
  PORTC|=0x01;//red only
   
  PORTC&=0xFE;
  PORTC|=0x04;//green only
  
  thisNodeAddr=(0xF0&PIND)>>4;
  if(thisNodeAddr==1)
  {
    nodeType = targetNode;
    nodeOrder = 0;
  }else{
    nodeType = baseStationNode;
    nodeOrder = thisNodeAddr-1;
  }
  LoRa.idle();
  LoRa.setPins(10, 9, 3);
  LoRa.setSignalBandwidth(master2NodeBW);
  LoRa.setSpreadingFactor(master2NodeSpreadingFactor);
  LoRa.setCodingRate4(master2NodeCodingRate);
  LoRa.setPreambleLength(master2NodePreambleLength);
  LoRa.setSyncWord(master2NodeSyncWord);
  LoRa.setTxPower(master2NodeTxPower);
  LoraLNASetting(master2NodeLNAStatus);
  LoRa.disableInvertIQ();
  LoRa.enableCrc();
  LoRa.onReceive(receiveIQR);
  LoRa.receive();
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  m2nTransmitTime = minTransmitTime*power(master2NodeSpreadingFactor-7)*(250E3/master2NodeBW);
  Serial.println("m2n: "+String(m2nTransmitTime));
}
