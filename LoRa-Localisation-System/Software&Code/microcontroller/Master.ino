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
const unsigned long robin2Duration = 2000;
const uint8_t maxNoOfNode = 12;
const uint8_t messagelength = 13;
const long minTransmitTime = 36;
const double master2NodeBW = 41.7E3;
const uint8_t master2NodeSpreadingFactor = 9;
const uint8_t master2NodeCodingRate = 8;
const unsigned short master2NodePreambleLength = 0x0F;
const uint8_t master2NodeSyncWord = 0xFF;
const uint8_t master2NodeTxPower = 17;
const uint8_t master2NodeLNAStatus = LNAAuto;
const uint8_t nTransmitionMode = 0;
const uint8_t cTransmitionMode = 1;
const uint8_t tTransmitionMode = 2;

const uint8_t masterNodeAddr = 0x00;
const uint8_t boardCastAddr = 0xFF;

bool newMessage = false;
String message = "";
String command;
String fullCommand;
unsigned short miniSecond;
unsigned long second;
uint8_t noOfAvailableNodes = 4;
uint8_t availableNodesAddr[maxNoOfNode] = {0};
unsigned long timeStamp;
bool roundRobinMode = false;
bool newSlaveTxTime;
bool commandAvailable = true;
unsigned short numberOfTransmitions;
unsigned long TransmitTill;
uint8_t transmitionMode;
double node2NodeBW = 41.7E3;
uint8_t node2NodeSpreadingFactor = 9;
long n2nTransmitTime;
long m2nTransmitTime;

//#####################################################################################################
//########################################################setup()######################################
//#####################################################################################################
void setup() 
{
  Serial.begin(9600);
  while(!Serial);
  Serial.println("Master side");
  int setUp;
  while(setUp = LoRa.begin(433E6))
  {
    if(setUp)
    {
      message.reserve(200);
      Serial.println(F("Connected to LoRa chips"));
      LoraMasterSlaveInitialization();
      timerInitialization();
      sei();
      Serial.println(F("Initialization completed"));
      break;
    }else{
      Serial.println(F("Failed to connect, retry in 10s"));
      delay(10E3);
    }
  }
}
//#####################################################################################################
//########################################################loop()#######################################
//#####################################################################################################
void loop() 
{
  if(newMessage)
  {
    Serial.println("Recieved Full Commmand: "+message);
    fullCommand = message;
    if(fullCommand.indexOf(' ')!=-1)
    {
      command = fullCommand.substring(0, fullCommand.indexOf(' '));
    }else{
      command = fullCommand;
    }
    message = "";
    newMessage = false;
  }
  if(roundRobinMode&&commandAvailable)
  {
    if(numberOfTransmitions)
    {
      if(second*1000+miniSecond>robin2Duration+timeStamp)
      {
        commandAvailable = false;
        Serial.println("Next");
        targetToNodeTransmit();
      }
    }else{
       roundRobinMode = false;
       Serial.println("End");
    }
  }
  if(commandAvailable)
  {
    if(command.equals("set"))
    {
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(4,fullCommand.indexOf(' ',4));
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      boolean nodeType = false;
      String type = fullCommand.substring(fullCommand.indexOf(' ',4)+1);
      if(type.equals("base"))
      {
        nodeType = false;
      }else if(type.equals("target")){
        nodeType = true;
      }else{
        Serial.println(F("node type unknown, default set to base station"));
      }
      setTypeForNode(nodeAddr,nodeType);
      command = "";
    }else if(command.equals("ping")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(5);
      nodeAddr = inputedAddr.toInt();
      pingNode(nodeAddr);
      command = "";
    }else if(command.equals("nodebw")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(7, fullCommand.indexOf(' ',7));
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      String InputedBW = fullCommand.substring(fullCommand.indexOf(' ',7)+1);
      unsigned long BW = 250E3;
      if(InputedBW.equals("250e3"))
      {
        BW = 250E3;
      }else if(InputedBW.equals("125e3")){
        BW = 125E3;
      }else if(InputedBW.equals("62.5e3")){
        BW = 62.5E3;
      }else if(InputedBW.equals("41.7e3")){
        BW = 41.7E3;
      }else if(InputedBW.equals("31.25e3")){
        BW = 31.25E3;
      }else if(InputedBW.equals("20.8e3")){
        BW = 20.8E3;
      }else if(InputedBW.equals("15.6e3")){
        BW = 15.6E3;
      }else if(InputedBW.equals("10.4e3")){
        BW = 10.4E3;
      }else if(InputedBW.equals("7.8e3")){
        BW = 7.8E3;
      }else{
        Serial.println(F("Invalid input BW, set to default 250E3"));
      }
      setBWForNode(nodeAddr, BW);
      command = "";
    }else if(command.equals("nodecr")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(7, fullCommand.indexOf(' ',7));
      uint8_t codingRate = (uint8_t)fullCommand.substring(fullCommand.indexOf(' ',7)+1).toInt();
      if(codingRate<5||codingRate>8)
      {
        Serial.println(F("Invalid input coding Rate, set to default 5"));
        codingRate = 5;
      }
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      setCRorNode(nodeAddr, codingRate);
      command = "";
    }else if(command.equals("nodesf")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(7, fullCommand.indexOf(' ',7));
      uint8_t spreadingFactor = (uint8_t)fullCommand.substring(fullCommand.indexOf(' ',7)+1).toInt();
      if(spreadingFactor<7||spreadingFactor>12)
      {
        Serial.println(F("Invalid input spreading factor, set to default 8"));
        spreadingFactor = 8;
      }
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      setSFForNode(nodeAddr, spreadingFactor);
      command = "";
    }else if(command.equals("nodetxpower")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(12, fullCommand.indexOf(' ',12));
      uint8_t txPower = (uint8_t)fullCommand.substring(fullCommand.indexOf(' ',12)+1).toInt();
      if(txPower<2||txPower>20)
      {
        Serial.println(F("Invalid input Transmition Power, set to default 20"));
        txPower = 20;
      }
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      setTxPowerForNode(nodeAddr, txPower);
      command = "";
    }else if(command.equals("nodelna")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(8, fullCommand.indexOf(' ',8));
      String inputLNAStatus = fullCommand.substring(fullCommand.indexOf(' ',8)+1);
      uint8_t LNAStatus = LNAAuto;
      if(inputLNAStatus.equals("auto"))
      {
        LNAStatus = LNAAuto;
      }else if(inputLNAStatus.equals("g1")){
        LNAStatus = LNAG1MAX;
      }else if(inputLNAStatus.equals("g2")){
        LNAStatus = LNAG2;
      }else if(inputLNAStatus.equals("g3")){
      LNAStatus = LNAG3;
      }else if(inputLNAStatus.equals("g4")){
        LNAStatus = LNAG4;
      }else if(inputLNAStatus.equals("g5")){
        LNAStatus = LNAG5;
      }else if(inputLNAStatus.equals("g6")){
        LNAStatus = LNAG6MIN;
      }else if(inputLNAStatus.equals("off")){
        LNAStatus = LNAG0;
      }else{
        Serial.println(F("Invalid input LNA Status, set to default Auto"));
      }
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      setLNAForNode(nodeAddr, LNAStatus);
      command = "";
    }else if(command.equals("timestampadjustment")){
      timeStampAdjust();
      command = "";
    }else if(command.equals("targettobasetransmitionn")){
      unsigned short numberOfTransmit = fullCommand.substring(25).toInt();
      if(numberOfTransmit)
      {
        targetToNodeTransmitN(numberOfTransmit);
      }else{
        Serial.println("Invalid Number input");
      }
      command = "";
    }else if(command.equals("targettobasetransmitionc")){
      targetToNodeTransmitC();
      command = "";
    }else if(command.equals("targettobasetransmitiont")){
      unsigned short secondOfTransmit = fullCommand.substring(25).toInt();
      if(secondOfTransmit)
      {
        targetToNodeTransmitT(secondOfTransmit);
      }else{
        Serial.println("Invalid Second input");
      }
      command = "";
    }else if(command.equals("discover")){
      uint8_t noOfNode = fullCommand.substring(9).toInt();
      nodeDiscovering(noOfNode);
      command = "";
    }else if(command.equals("nodepl")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(7, fullCommand.indexOf(' ',7));
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      unsigned short preambleLength = fullCommand.substring(fullCommand.indexOf(' ',7)+1).toInt();
      setPLForNode(nodeAddr, preambleLength);
      command = "";
    }else if(command.equals("nodesw")){
      uint8_t nodeAddr;
      String inputedAddr = fullCommand.substring(7, fullCommand.indexOf(' ',7));
      if(inputedAddr.equals("ff"))
      {
        nodeAddr = boardCastAddr;
      }else{
        nodeAddr = inputedAddr.toInt();
      }
      uint8_t syncWord = fullCommand.substring(fullCommand.indexOf(' ',7)+1).toInt();
      setSWForNode(nodeAddr, syncWord);
      command = "";
    }else if(command.equals("break")){
      Serial.println(F("Break for Target to Base Transmition!"));
      numberOfTransmitions = 0;
      roundRobinMode = false;
      Serial.println("End");
      command = "";
    }
  }else{
    if((second*1000+miniSecond)>(timeStamp+5000+n2nTransmitTime+50+(m2nTransmitTime+50)*(noOfAvailableNodes-1)))
    {
      LoRa.onReceive(false);
      switchToTransmit();
      Serial.println(F("Timeout for Base Transmition occurs!"));
      if(transmitionMode==nTransmitionMode)
      {
        numberOfTransmitions--;
      }else if(transmitionMode==tTransmitionMode){
        if(TransmitTill<second*1000+miniSecond)
        {
          Serial.println(F("Timelimit reached!"));
          numberOfTransmitions = 0;
        }
      }
      timeStamp = second*1000+miniSecond;
      commandAvailable = true;
    }
  }
}
//#####################################################################################################
//########################################################function#####################################
//#####################################################################################################
//target To Node Transmit N////////////////////////////////////////////////////////////////////////////
void targetToNodeTransmitN(unsigned short numberOfTransmit)
{
  sei();
  transmitionMode = nTransmitionMode;
  numberOfTransmitions = numberOfTransmit;
  newSlaveTxTime = true;
  commandAvailable = false;
  Serial.println("No. of Transmition: "+String(numberOfTransmitions));
  roundRobinMode = true;
  targetToNodeTransmit();
}
//target To Node Transmit Continuous////////////////////////////////////////////////////////////////////////////
void targetToNodeTransmitC()
{
  sei();
  transmitionMode = cTransmitionMode;
  numberOfTransmitions = 1;
  newSlaveTxTime = true;
  commandAvailable = false;
  Serial.println("Continuous Transmition, input \"break\" to end");
  roundRobinMode = true;
  targetToNodeTransmit();
}
//target To Node Transmit Timelimit////////////////////////////////////////////////////////////////////////////
void targetToNodeTransmitT(unsigned short timelimitSecond)
{
  sei();
  transmitionMode = tTransmitionMode;
  numberOfTransmitions = 1;
  newSlaveTxTime = true;
  commandAvailable = false;
  Serial.println("Timelimit Transmition for "+String(timelimitSecond)+" Second(s)");
  roundRobinMode = true;
  TransmitTill = timelimitSecond*1000+second*1000+miniSecond;
  targetToNodeTransmit();
}
//target To Node Transmit//////////////////////////////////////////////////////////////////////////////
void targetToNodeTransmit()
{
  sei();
  String mMessage = "";
  mMessage+=(char)boardCastAddr;
  mMessage+=(char)masterNodeAddr;
  mMessage+=(char)startTargetToBaseTransmition;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  timeStamp = second*1000+miniSecond;
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  switchToRecieve();
  LoRa.onReceive(receiveIQR);
}
//switch To Recieve////////////////////////////////////////////////////////////////////////////////////
void switchToRecieve()
{
  LoRa.idle();
  LoRa.enableInvertIQ();
  LoRa.receive();
}
//switch To Transmit///////////////////////////////////////////////////////////////////////////////////
void switchToTransmit()
{
  LoRa.idle();
  LoRa.disableInvertIQ();
  LoRa.receive();
}
//Discover nodes/////////////////////////////////////////////////////////////////////////////////////
void nodeDiscovering(uint8_t noOfNode)
{
  sei();
  noOfAvailableNodes = 0;
  if(noOfNode==0)
  {
    for(uint8_t nodeAddress = 1;nodeAddress<=maxNoOfNode;nodeAddress++)
    {
      if(pingNode(nodeAddress, m2nTransmitTime+22))
      {
        nodeOrder(nodeAddress, noOfAvailableNodes);
        availableNodesAddr[noOfAvailableNodes] = nodeAddress;
        nodeTurn(nodeAddress, nodeOn);
        Serial.println("Set Node address "+String(nodeAddress)+" to order "+String(noOfAvailableNodes));
        noOfAvailableNodes++;
        if(noOfAvailableNodes==1)
        {
          setTypeForNode(nodeAddress, targetNode);
        }else{
          setTypeForNode(nodeAddress, baseStationNode);
        }
      }
    }
  }else{
    for(uint8_t nodeAddress = 1;nodeAddress<=maxNoOfNode;nodeAddress++)
    {
      if(noOfAvailableNodes<noOfNode)
      {
        if(pingNode(nodeAddress, 250))
        {
          nodeOrder(nodeAddress, noOfAvailableNodes);
          availableNodesAddr[noOfAvailableNodes] = nodeAddress;
          nodeTurn(nodeAddress, nodeOn);
          Serial.println("Set Node address "+String(nodeAddress)+" to order "+String(noOfAvailableNodes));
          noOfAvailableNodes++;
          if(noOfAvailableNodes==1)
          {
            setTypeForNode(nodeAddress, targetNode);
          }else{
            setTypeForNode(nodeAddress, baseStationNode);
          }
        }
      }else{
        nodeTurn(nodeAddress, nodeOff);
      }
    }
  }
  Serial.print("Available Nodes: ");
  for(int nodeIndex = 0;nodeIndex<noOfAvailableNodes ;nodeIndex++)
  {
    Serial.print(String(availableNodesAddr[nodeIndex])+", ");
  }
  Serial.println();
  if(noOfAvailableNodes<maxNoOfNode)
  {
    availableNodesAddr[noOfAvailableNodes] = 0;
  }
  Serial.println("Total number of "+String(noOfAvailableNodes)+" node(s) discovered");
}
//enable/disable node/////////////////////////////////////////////////////////////////////////////////
void nodeTurn(uint8_t nodeAddress,boolean nodeOff)
{
  sei();
  String mMessage = "";
  mMessage+=(char)nodeAddress;
  mMessage+=(char)masterNodeAddr;
  mMessage+=(char)nodeOnOff;
  mMessage+=(char)nodeOff;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//send node order/////////////////////////////////////////////////////////////////////////////////////
void nodeOrder(uint8_t nodeAddr,uint8_t order)
{
  sei();
  String mMessage = "";
  mMessage+=(char)nodeAddr;
  mMessage+=(char)masterNodeAddr;
  mMessage+=(char)setNodeOrder;
  mMessage+=(char)order;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//Adjust timeStamp For Node/////////////////////////////////////////////////////////////////////////////////////
void timeStampAdjust()
{
  sei();
  String mMessage = "";
  mMessage+=(char)boardCastAddr;
  mMessage+=(char)masterNodeAddr;
  mMessage+=(char)timeStampAdjustment;
  while(mMessage.length()<(messagelength-4))
  {
    mMessage+='~';
  }
  unsigned long timeStamp = second*1000+miniSecond;
  uint8_t timeStamps[4];
  timeStamps[0] = 0xFF&(timeStamp>>24);
  timeStamps[1] = 0xFF&(timeStamp>>16);
  timeStamps[2] = 0xFF&(timeStamp>>8);
  timeStamps[3] = 0xFF&(timeStamp);
  mMessage+=(char)timeStamps[0];
  mMessage+=(char)timeStamps[1];
  mMessage+=(char)timeStamps[2];
  mMessage+=(char)timeStamps[3];
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  /*
  Serial.println("bit1 "+String((uint8_t)mMessage.charAt(9)));
  Serial.println("bit2 "+String((uint8_t)mMessage.charAt(10)));
  Serial.println("bit3 "+String((uint8_t)mMessage.charAt(11)));
  Serial.println("bit4 "+String((uint8_t)mMessage.charAt(12)));
  
  Serial.println("lens: "+String(mMessage.length()));
  Serial.println("Master1 ts: "+String(timeStamp));
  Serial.println("bit1 "+String((uint8_t)mMessage.charAt(9)));
  Serial.println("bit2 "+String((uint8_t)mMessage.charAt(10)));
  Serial.println("bit3 "+String((uint8_t)mMessage.charAt(11)));
  Serial.println("bit4 "+String((uint8_t)mMessage.charAt(12)));
  
  char c = 0xF0;
  unsigned long cc = (uint8_t)c;
  Serial.println("cc1 ts: "+String(cc));

  cc = (unsigned short)c<<8;
  Serial.println("cc2 ts: "+String(cc));

  cc = (unsigned long)((unsigned short)c<<8)<<8;
  Serial.println("cc3 ts: "+String(cc));

  cc = (unsigned long)c<<24;
  Serial.println("cc4 ts: "+String(cc));
  unsigned long reconstructedTimeStamp = (uint8_t)mMessage.charAt(12)|(unsigned short)mMessage.charAt(11)<<8|(unsigned long)((unsigned short)mMessage.charAt(10)<<8)<<8|(unsigned long)mMessage.charAt(9)<<24;
  Serial.println("Master2 ts: "+String(reconstructedTimeStamp));
  */
  Serial.println("Time sent to nodes is: "+String(timeStamp));
}
//set LNA For Node/////////////////////////////////////////////////////////////////////////////////////
void setLNAForNode(uint8_t nodeAddr,uint8_t LNAStatus)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeLNA;
  mMessage += (char)LNAStatus;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//set Tx Power For Node/////////////////////////////////////////////////////////////////////////////////////
void setTxPowerForNode(uint8_t nodeAddr, uint8_t txPower)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeTxPower;
  mMessage += (char)txPower;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//set Sprending Factor For Node/////////////////////////////////////////////////////////////////////////////////////
void setSFForNode(uint8_t nodeAddr, uint8_t spreadingFactor)
{
  sei();
  node2NodeSpreadingFactor = spreadingFactor;
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeSF;
  mMessage += (char)spreadingFactor;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  newSlaveTxTime = true;
}
//set CR ForvNode/////////////////////////////////////////////////////////////////////////////////////
void setCRorNode(uint8_t nodeAddr, uint8_t codingRate)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeCR;
  mMessage += (char)codingRate;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  newSlaveTxTime = true;
}
//set Bandwidth For Node/////////////////////////////////////////////////////////////////////////////////////
void setBWForNode(uint8_t nodeAddr, unsigned long BW)
{
  sei();
  node2NodeBW = (double)BW;
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeBW;
  switch(BW)
  {
    case((long)250E3):
    {
      mMessage += (char)BW250E3;
      break;
    }
    case((long)125E3):
    {
      mMessage += (char)BW125E3;
      break;
    }
    case((long)62.5E3):
    {
      mMessage += (char)BW62P5E3;
      break;
    }
    case((long)41.7E3):
    {
      mMessage += (char)BW41P7E3;
      break;
    }
    case((long)31.25E3):
    {
      mMessage += (char)BW31P25E3;
      break;
    }
    case((long)20.8E3):
    {
      mMessage += (char)BW20P8E3;
      break;
    }
    case((long)15.6E3):
    {
      mMessage += (char)BW15P6E3;
      break;
    }
    case((long)10.4E3):
    {
      mMessage += (char)BW10P4E3;
      break;
    }
    case((long)7.8E3):
    {
      mMessage += (char)BW7P8E3;
      break;
    }
  }
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  unsigned long startTime = second*1000+miniSecond;
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  Serial.println("send in "+String(second*1000+miniSecond - startTime)+"ms");
  newSlaveTxTime = true;
}
//set Preamble Length For Node/////////////////////////////////////////////////////////////////////////////////////
void setPLForNode(uint8_t nodeAddr, unsigned short preambleLength)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodePL;
  uint8_t preambleLengths[2];
  preambleLengths[0] = 0xFF&(preambleLength>>8);
  preambleLengths[1] = 0xFF&preambleLength;
  mMessage += (char)preambleLengths[0];
  mMessage += (char)preambleLengths[1];
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  newSlaveTxTime = true;
}
//set sync Word For Node/////////////////////////////////////////////////////////////////////////////////////
void setSWForNode(uint8_t nodeAddr, uint8_t syncWord)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeSW;
  mMessage += (char)syncWord;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//set Type For Node/////////////////////////////////////////////////////////////////////////////////////
void setTypeForNode(uint8_t nodeAddr, boolean nodeType)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)setNodeType;
  mMessage += (char)nodeType;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
}
//ping Node 1/////////////////////////////////////////////////////////////////////////////////////
bool pingNode(uint8_t nodeAddr)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)ping;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  Serial.println("Pinging Node"+String(nodeAddr));
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  unsigned long startTime = second*1000+miniSecond;
  switchToRecieve();
  String recievedMessage = pollForRecieve(2.5*m2nTransmitTime);
  switchToTransmit();
  if(recievedMessage.length()==0)
  {
    Serial.println("Ping failed1");
    return false;
  }else{
    if(recievedMessage.charAt(0)==masterNodeAddr&&recievedMessage.charAt(1)==nodeAddr)
    {
      short forwardRssi = (recievedMessage.charAt(3)<<8)|(recievedMessage.charAt(4));
      short backwardRssi = LoRa.packetRssi();
      Serial.println("Node"+String(nodeAddr)+" responded in "+String(second*1000+miniSecond - startTime)+"ms RSSI: "+String(forwardRssi)+", "+String(backwardRssi));
      return true;
    }else{
      Serial.println("Ping failed2");
      return false;
    }
  }
  
}
//ping Node 2/////////////////////////////////////////////////////////////////////////////////////
bool pingNode(uint8_t nodeAddr, unsigned long timeout)
{
  sei();
  String mMessage = "";
  mMessage += (char)nodeAddr;
  mMessage += (char)masterNodeAddr;
  mMessage += (char)ping;
  while(mMessage.length()<messagelength)
  {
    mMessage+='~';
  }
  Serial.println("Pinging Node"+String(nodeAddr));
  LoRa.beginPacket();
  LoRa.print(mMessage);
  LoRa.endPacket();
  unsigned long startTime = second*1000+miniSecond;
  switchToRecieve();
  String recievedMessage = pollForRecieve(timeout);
  switchToTransmit();
  if(recievedMessage.length()==0)
  {
    Serial.println("Ping failed1");
    return false;
  }else{
    if(recievedMessage.charAt(0)==masterNodeAddr&&recievedMessage.charAt(1)==nodeAddr)
    {
      short forwardRssi = (recievedMessage.charAt(3)<<8)|(recievedMessage.charAt(4));
      short backwardRssi = LoRa.packetRssi();
      Serial.println("Node"+String(nodeAddr)+" responded in "+String(second*1000+miniSecond - startTime)+"ms RSSI: "+String(forwardRssi)+", "+String(backwardRssi));
      return true;
    }else{
      Serial.println("Ping failed2");
      return false;
    }
  }
}
//poll For Recieve/////////////////////////////////////////////////////////////////////////////////////
String pollForRecieve(unsigned long timeout)
{
  unsigned long endTime = second*1000+miniSecond+timeout;
  String recievedMessage = "";
  while(!LoRa.parsePacket())
  {
    if(second*1000+miniSecond == endTime)
    {
      return "";
    }
  }
  while(LoRa.available())
  {
    recievedMessage += (char)LoRa.read();
  }
  return recievedMessage;
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
//Dio0 RX Interrupt/////////////////////////////////////////////////////////////////////////////////////
void receiveIQR(int size)
{
  String messageFromBase = "";
  while(LoRa.available())
  {
    messageFromBase += (char)LoRa.read();
  }
  if((uint8_t)messageFromBase.charAt(0)==masterNodeAddr&&(uint8_t)messageFromBase.charAt(2)==baseToMasterTransmition)
  {
    if(newSlaveTxTime)
    {
      Serial.println("N2N:"+String(second*1000+miniSecond-timeStamp)+"ms");
      newSlaveTxTime = false;
    }
    uint8_t messageFromBaseArray[10];
    for(int mesageIndex = 3;mesageIndex<13;mesageIndex++)
    {
      messageFromBaseArray[mesageIndex-3] = (uint8_t)messageFromBase.charAt(mesageIndex);
    }
    //short forwardRssi = (unsigned short)(messageFromBaseArray[0]<<8)|(uint8_t)(messageFromBaseArray[1]);
    short PKR = (uint8_t)messageFromBaseArray[0]-164;
    float SNR = ((int8_t)messageFromBaseArray[1])*0.25;
    signed long fError = (uint8_t)messageFromBaseArray[5]|(unsigned short)messageFromBaseArray[4]<<8|(unsigned long)((unsigned short)messageFromBaseArray[3]<<8)<<8|(unsigned long)messageFromBaseArray[2]<<24;
    short Alpha = (uint8_t)messageFromBaseArray[6]-164;
    Serial.println("Base addr:"+String((uint8_t)messageFromBase.charAt(1))+" T2B PKR:"+String(PKR)+" SNR:"+String(SNR)+" fEr:"+String(fError)+" Alp:"+String(Alpha)+" FS");
    if(availableNodesAddr[noOfAvailableNodes-1]==(uint8_t)messageFromBase.charAt(1))
    {
      if(transmitionMode==nTransmitionMode)
      {
        numberOfTransmitions--;
      }else if(transmitionMode==tTransmitionMode){
        if(TransmitTill<second*1000+miniSecond)
        {
          Serial.println(F("Timelimit reached!"));
          numberOfTransmitions = 0;
        }
      }
      timeStamp = second*1000+miniSecond;
      commandAvailable = true;
      switchToTransmit();
      LoRa.onReceive(false);
    }
  }
}
//serial after-loop() semi-Interrupt////////////////////////////////////////////////////////////////////////////////////
void serialEvent()
{
  while(Serial.available())
  {
    byte chara = Serial.read();
    if(chara==10)
    {
      newMessage = true;
      continue;
    }
    if(chara==13)
    {
      continue;
    }
    message+=(char)chara;
  }
}
//#####################################################################################################
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
void LoraMasterSlaveInitialization()
{
  LoRa.idle();
  LoRa.setSignalBandwidth(master2NodeBW);
  LoRa.setSpreadingFactor(master2NodeSpreadingFactor);
  LoRa.setCodingRate4(master2NodeCodingRate);
  LoRa.setPreambleLength(master2NodePreambleLength);
  LoRa.setSyncWord(master2NodeSyncWord);
  LoRa.setTxPower(master2NodeTxPower);
  LoRa.disableInvertIQ();
  LoRa.enableCrc();
  n2nTransmitTime = minTransmitTime*power(node2NodeSpreadingFactor-7)*(250E3/node2NodeBW);
  Serial.println("n2n: "+String(n2nTransmitTime));
  m2nTransmitTime = minTransmitTime*power(master2NodeSpreadingFactor-7)*(250E3/master2NodeBW);
  Serial.println("m2n: "+String(m2nTransmitTime));
}
