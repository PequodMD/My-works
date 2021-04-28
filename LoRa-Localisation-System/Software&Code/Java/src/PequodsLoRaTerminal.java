package localizationLoRa;

import java.awt.GridBagConstraints;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.text.DefaultCaret;

import org.apache.commons.math3.ml.clustering.DBSCANClusterer;
import org.apache.commons.math3.ml.clustering.KMeansPlusPlusClusterer;
import org.apache.commons.math3.ml.clustering.MultiKMeansPlusPlusClusterer;
import org.apache.commons.math3.util.CombinatoricsUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gnu.io.CommPort;

public class PequodsLoRaTerminal 
{
	private InputStream serialPortInputStream;
	private OutputStream serialPortOutputStream;
	private CommPort serialPort;
	private StringBuilder content;
	private final byte idle = 0;
	private final byte exit = 1;
	private final byte end = 2;
	private final byte raw = 3;
	private final byte measureAlpha = 4;
	private final byte linealityTest = 5;
	private final byte locateTarget = 6;
	private final byte measureAlphaNlocateTarget = 7;
	private final byte measureAlphaNlocateTargetMirror = 8;
	
	public static String input;
	public static boolean newInput;
	
	public PequodsLoRaTerminal(List<Object> inOutputStream)
	{
		serialPortInputStream = (InputStream)inOutputStream.get(0);
		serialPortOutputStream = (OutputStream)inOutputStream.get(1);
		serialPort = (CommPort)inOutputStream.get(2);
	}
	public boolean start(Scanner scanner)throws IOException
	{
		short numberOfSample = 100;
		List<Integer> availableNode = new ArrayList<Integer>();
		List<Double> RSSI = new ArrayList<Double>();
		double tempDistance = 0.0;
		List<Double> distance = new ArrayList<Double>();
		List<List<Double>> RSSIList = new ArrayList<List<Double>>();
		List<List<Double>> AlphaRSSIList = new ArrayList<List<Double>>();
		
		double alpha = 0;
		double RSSI1m = 0;
		List<Double> baseX = new ArrayList<Double>();
		List<Double> baseY = new ArrayList<Double>();
		double MasterX = 0;
		double MasterY = 0;
		double TargetX = 0;
		double TargetY = 0;
		int initNum = 0;
		
		boolean mirrorFlag = false;
		boolean toExit = false;
		File textLog;
		SimpleDateFormat fileNameExtension = new SimpleDateFormat("yyyyMMMdd_HHmmss");
		File basedir = new File("C:\\LoRa");
		if(!basedir.isDirectory())
		{
			basedir.mkdirs();
		}
    	textLog = new File(basedir,  "log"+fileNameExtension.format(new Date())+".txt");
    	content = new StringBuilder();

    	input = new String();
    	newInput = false;
    	List<String> inputToken = new ArrayList<String>();
    	byte currentOperation = idle;
    	byte currentOperationStage = 0;
    	Thread inputReader = new Thread(new SerialReader(scanner));
    	inputReader.start();
    	
    	StringBuilder SerialPortInputTransmitionBuffer = new StringBuilder();
    	List<String> SerialPortInputTransmition = new ArrayList<String>();
    	
		byte[] buffer = new byte[1024];
		int len = -1;
		JFrame f = new JFrame();
		f.setTitle(serialPort.getName()+" Monitor");
		
		JPanel jp = new JPanel();
		
		JTextArea panel = new JTextArea(20 ,40);
		panel.setEditable(false);
		DefaultCaret caret = (DefaultCaret)panel.getCaret();
		caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
		JScrollPane scollPane = new JScrollPane(panel);
		
		GridBagConstraints c = new GridBagConstraints();
		c.gridwidth = GridBagConstraints.REMAINDER;
		c.fill = GridBagConstraints.BOTH;
		c.weightx = 1.0;
		c.weighty = 1.0;
		jp.add(scollPane,c);
		
		f.add(jp);
		f.pack();
		f.setVisible(true);
		
		terminalLoop:
		while(true) 
		{
			if((len = serialPortInputStream.read(buffer))>-1)
			{
				String inString = new String(buffer,0,len);
				if(!inString.isEmpty())
				{
					String[] inStrings = inString.split("\n");
					for(int i=0;i<inStrings.length;i++)
					{
						if(i>0)
						{
							String fullSingleInputTransmition = SerialPortInputTransmitionBuffer.toString()+"\n";
							SerialPortInputTransmition.add(fullSingleInputTransmition);
							content.append(">>"+fullSingleInputTransmition);
							panel.append(fullSingleInputTransmition);
							SerialPortInputTransmitionBuffer = new StringBuilder("");
						}
						SerialPortInputTransmitionBuffer.append(inStrings[i]);
					}
					if(inString.endsWith("\n"))
					{
						String fullSingleInputTransmition = SerialPortInputTransmitionBuffer.toString()+"\n";
						SerialPortInputTransmition.add(fullSingleInputTransmition);
						content.append(">>"+fullSingleInputTransmition);
						panel.append(fullSingleInputTransmition);
						SerialPortInputTransmitionBuffer = new StringBuilder("");
					}
				}
				
			}
			if(newInput)
			{
				content.append("<<"+input+'\n');
				inputToken = Arrays.asList(input.split(" "));
				input = new String("");
				newInput = false;
				
				StringBuilder userInput = new StringBuilder();
				userInput.append("Received:");
				Iterator<String> commandIterator = inputToken.iterator();
				while(commandIterator.hasNext())
				{
					userInput.append(" "+commandIterator.next());
				}
				userInput.append('\n');
				systemOut(userInput.toString());
			}
			if(!inputToken.isEmpty())
			{
				if(inputToken.get(0).equalsIgnoreCase("break"))
				{
					currentOperation = idle;
					currentOperationStage = 0;
					inputToken = new ArrayList<String>();
				}
			}
			switch(currentOperation)
			{
				case(idle):
				{
					SerialPortInputTransmition = new ArrayList<String>();
					if(!inputToken.isEmpty())
					{
						if(inputToken.get(0).equalsIgnoreCase("exit"))
						{
							currentOperation = exit;
						}else if(inputToken.get(0).equalsIgnoreCase("end")){
							currentOperation = end;
						}else if(inputToken.get(0).equalsIgnoreCase("raw")&&inputToken.size()>1) {
							currentOperation = raw;
						}else if(inputToken.get(0).equalsIgnoreCase("measureAlpha")&&inputToken.size()>1) {
							currentOperation = measureAlpha;
						}else if(inputToken.get(0).equalsIgnoreCase("linealityTest")) {
							currentOperation = linealityTest;
						}else if(inputToken.get(0).equalsIgnoreCase("locateTarget")&&inputToken.size()>1) {
							currentOperation = locateTarget;
						}else if(inputToken.get(0).equalsIgnoreCase("locateTargetNAlpha")&&inputToken.size()>1) {
							currentOperation = measureAlphaNlocateTarget;
						}else if(inputToken.get(0).equalsIgnoreCase("measureAlphaNTargetMi")&&inputToken.size()>1) {
							currentOperation = measureAlphaNlocateTargetMirror;
						}else{
							systemOut("Unknown command/Incorrect use of command\n");
							inputToken = new ArrayList<String>();
						}
					}
					break;
				}
				case(end):
				{
					toExit = false;
					break terminalLoop;
				}
				case(exit):
				{
					toExit = true;
					break terminalLoop;
				}
				case(raw):
				{
					Iterator<String> commands = inputToken.iterator();
					commands.next();
					StringBuilder rawCommand = new StringBuilder(commands.next().toLowerCase());
					while(commands.hasNext())
					{
						rawCommand.append(" "+commands.next().toLowerCase());
					}
					sendToSerial(rawCommand.toString()+'\n');
					currentOperation = idle;
					inputToken = new ArrayList<String>();
					break;
				}
				case(measureAlpha):
				{
					switch(currentOperationStage)
					{
						case(0):
						{
							SerialPortInputTransmition = new ArrayList<String>();
							sendToSerial("timestampadjustment\n");
							try 
			    			{
			    				TimeUnit.SECONDS.sleep(1);
			    			} catch (InterruptedException e){
			    				e.printStackTrace();
			    			}
							sendToSerial("discover 2\n");
							numberOfSample = (short)Integer.parseInt(inputToken.get(1));
							systemOut("Measure Alpha with "+numberOfSample+" sample\n");
							currentOperationStage = 1;
							break;
						}
						case(1):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("Total number of "))
								{
									if(Integer.parseInt(line.substring(16, 17))>1)
									{
										systemOut("Input ok when ready to measure 1m RSSI\n");
										currentOperationStage = 2;
									}else{
										systemOut("Not enough available node! Return to idle\n");
										currentOperationStage = 0;
										currentOperation = idle;
									}
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(2):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.get(0).equalsIgnoreCase("ok"));
								{
									currentOperationStage = 3;
									RSSIList = new ArrayList<List<Double>>();
									RSSI = new ArrayList<Double>();
									inputToken = new ArrayList<String>();
									sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
								}
							}
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(3):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("PKR:"))
								{
									RSSI.add(Double.parseDouble(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
								}
								if(line.contains("End"))
								{
									systemOut("Input ok when ready to measure 10m RSSI\n");
									RSSIList.add(RSSI);
									RSSI = new ArrayList<Double>();
									currentOperationStage = 4;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(4):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.get(0).equalsIgnoreCase("ok"));
								{
									currentOperationStage = 5;
									inputToken = new ArrayList<String>();
									sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
								}
							}
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(5):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("PKR:"))
								{
									RSSI.add(Double.parseDouble(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
								}
								if(line.contains("End"))
								{
									RSSI1m = PequodsMaths.getMean(RSSIList.get(0));
									double RSSI1Mean = PequodsMaths.getMean(RSSI);
									alpha = PequodsMaths.calculateAlpha(RSSI1m, RSSI1Mean);
									systemOut("RSSI 1m mean: "+RSSI1m+'\n');
									systemOut("RSSI 10m mean: "+RSSI1Mean+'\n');
									systemOut("Alpha: "+alpha+'\n');
									currentOperationStage = 0;
									currentOperation = idle;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						default:
						{
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
					}
					break;
				}
				case(linealityTest):
				{
					switch(currentOperationStage)
					{
						case(0):
						{
							SerialPortInputTransmition = new ArrayList<String>();
							sendToSerial("timestampadjustment\n");
							try 
			    			{
			    				TimeUnit.SECONDS.sleep(1);
			    			} catch (InterruptedException e){
			    				e.printStackTrace();
			    			}
							sendToSerial("discover\n");
							systemOut("lineality Test\n");
							currentOperationStage = 1;
							break;
						}
						case(1):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("Total number of "))
								{
									if(Integer.parseInt(line.substring(16, 17))>1)
									{
										systemOut("Test starts now\n");
										RSSI = new ArrayList<Double>();
										distance = new ArrayList<Double>();
										currentOperationStage = 2;
									}else {
										systemOut("Not enough available node! Return to idle\n");
										currentOperationStage = 0;
										currentOperation = idle;
									}
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(2):
						{
							systemOut("Input distance(m) if ready to measure RSSI, Input check if it is enough\n");
							currentOperationStage = 3;
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(3):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.get(0).equalsIgnoreCase("check"))
								{
									currentOperationStage = 5;
								}else {
									tempDistance = Double.parseDouble(inputToken.get(0));
									currentOperationStage = 4;
									sendToSerial("targettobasetransmitionn "+1+"\n");
									inputToken = new ArrayList<String>();
								}
							}
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(4):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("PKR:"))
								{
									RSSI.add(Double.parseDouble(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
									distance.add(new Double(tempDistance));
								}
								if(line.contains("End"))
								{
									currentOperationStage = 2;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(5):
						{
							systemOut("Start calculation\n");
							File outputxlsxFile = new File(basedir.getAbsolutePath()+fileNameExtension.format(new Date())+"Book.xlsx");
							XSSFWorkbook outputworkbook = new XSSFWorkbook();
							POIXMLProperties xmlProps = outputworkbook.getProperties();    
							POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
							coreProps.setCreator("Pequod");
							XSSFSheet rawDataSheet = outputworkbook.createSheet("RawData");
							int rawDataSheetCurrentRow = 0;
							
							XSSFRow tagRow = rawDataSheet.createRow(0);
							tagRow.createCell(0, CellType.STRING).setCellValue("Distance");
							tagRow.createCell(1, CellType.STRING).setCellValue("RSSI");
							rawDataSheetCurrentRow++;
							
							for(int index = 0;index<distance.size();index++)//cal
							{
								XSSFRow recordRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
								recordRow.createCell(0, CellType.NUMERIC).setCellValue(distance.get(index));
								recordRow.createCell(1, CellType.NUMERIC).setCellValue(RSSI.get(index));
								systemOut("Distance: "+distance.get(index)+", mean RSSI: "+RSSI.get(index)+"\n");
								rawDataSheetCurrentRow++;
							}
							double[] result = PequodsMaths.LinealityTest(distance, RSSI);
							systemOut("Slope: "+result[0]+'\n');
							systemOut("R square: "+result[1]+'\n');
							FileOutputStream writeFile = new FileOutputStream(outputxlsxFile);
							outputworkbook.write(writeFile);
							outputworkbook.close();
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							currentOperationStage = 0;
							currentOperation = idle;
							break;
						}
						default:
						{
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
					}
					break;
				}
				case(locateTarget):
				{
					switch(currentOperationStage)
					{
						case(0):
						{
							if(alpha==0)
							{
								systemOut("Can't locate target without a measured Alpha\n");
								currentOperationStage = 0;
								currentOperation = idle;
								inputToken = new ArrayList<String>();
								SerialPortInputTransmition = new ArrayList<String>();
							}else{
								sendToSerial("discover\n");
								numberOfSample = (short)Integer.parseInt(inputToken.get(1));
								systemOut("locate target with "+numberOfSample+" sample\n");
								availableNode = new ArrayList<Integer>();
								baseX = new ArrayList<Double>();
								baseY = new ArrayList<Double>();
								currentOperationStage = 1;
							}
							break;
						}
						case(1):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("Available Nodes: "))
								{
									String[] nodeAddress = line.substring(line.indexOf(", ")+2,line.length()-2).split(", ");
									for(int index = 0;index<nodeAddress.length;index++)
									{
										availableNode.add(Integer.parseInt(nodeAddress[index]));
									}
								}
								if(line.contains("Total number of "))
								{
									if(Integer.parseInt(line.substring(16, 17))>3)
									{
										RSSIList = new ArrayList<List<Double>>(availableNode.size());
										for(int index = 0;index<availableNode.size();index++)
										{
											RSSIList.add(new ArrayList<Double>());
										}
										currentOperationStage = 2;
									}else{
										systemOut("Not enough available node! Return to idle\n");
										currentOperationStage = 0;
										currentOperation = idle;
									}
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(2):
						{
							systemOut("Input the x, y coordinate for BaseStation address: "+availableNode.get(baseX.size())+"\n");
							SerialPortInputTransmition = new ArrayList<String>();
							inputToken = new ArrayList<String>();
							currentOperationStage = 3;
							break;
						}
						case(3):
						{
							if(!inputToken.isEmpty())
							{
								baseX.add(new Double(inputToken.get(0)));
								baseY.add(new Double(inputToken.get(1)));
								if(baseX.size()==availableNode.size())
								{
									currentOperationStage = 4;
									SerialPortInputTransmition = new ArrayList<String>();
									sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
								}else {
									currentOperationStage = 2;
								}
							}
							inputToken = new ArrayList<String>();
							break;
						}
						case(4):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("T2B PKR:"))
								{
									int nodeAddress = Integer.parseInt(line.substring(line.indexOf("addr:")+5,line.indexOf(" T2B")));
									RSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
								}
								if(line.contains("End"))
								{
									systemOut("Start calculation\n");
									List<Double> RSSImean = new ArrayList<Double>();
									for(int index = 0;index<availableNode.size();index++)
									{
										double mean = PequodsMaths.getMean(RSSIList.get(index));
										RSSImean.add(mean);
										systemOut("node address:"+availableNode.get(index)+" x:"+baseX.get(index)+" y:"+baseY.get(index)+" RSSI:"+mean+"\n");
									}
									systemOut("1m RSSI: "+RSSI1m+"\n");
									systemOut("Alpha: "+alpha+"\n");
									double[] result = PequodsMaths.LLSForLocalization(RSSImean, baseX, baseY, RSSI1m, alpha);
									systemOut("x: "+result[0]+"\n");
									systemOut("y: "+result[1]+"\n");
									systemOut("x^2+y^2: "+result[2]+"\n");
									systemOut("R square: "+result[3]+"\n");
									currentOperationStage = 0;
									currentOperation = idle;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						default:
						{
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
					}
				}
				case(measureAlphaNlocateTarget):
				{
					switch(currentOperationStage)
					{
						case(0):
						{
							sendToSerial("discover\n");
							numberOfSample = (short)Integer.parseInt(inputToken.get(1));
							systemOut("locate target with "+numberOfSample+" sample\n");
							availableNode = new ArrayList<Integer>();
							baseX = new ArrayList<Double>();
							baseY = new ArrayList<Double>();
							currentOperationStage = 1;
							break;
						}
						case(1):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("Available Nodes: "))
								{
									String[] nodeAddress = line.substring(line.indexOf(", ")+2,line.length()-2).split(", ");
									for(int index = 0;index<nodeAddress.length;index++)
									{
										if(Integer.parseInt(nodeAddress[index])==2)
										{
											systemOut("Excluded node2\n");
										}else {
											availableNode.add(Integer.parseInt(nodeAddress[index]));
										}
									}
								}
								if(line.contains("Total number of "))
								{
									//!!!!!!!!!!!!!change is if no. of node>10
									if(availableNode.size()>3)
									{
										RSSIList = new ArrayList<List<Double>>(availableNode.size());
										AlphaRSSIList = new ArrayList<List<Double>>();
										for(int index = 0;index<availableNode.size();index++)
										{
											RSSIList.add(new ArrayList<Double>());
											AlphaRSSIList.add(new ArrayList<Double>());
										}
										currentOperationStage = 2;
									}else{
										systemOut("Not enough available node! Return to idle\n");
										currentOperationStage = 0;
										currentOperation = idle;
									}
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(2):
						{
							StringBuilder stringOut = new StringBuilder("Master Node");
							for(int dummy = 0;dummy<availableNode.size();dummy++)
							{
								stringOut.append(","+availableNode.get(dummy).toString());
							}
							systemOut("Input the x, y coordinate for BaseStation address: "+stringOut+"\n");
							SerialPortInputTransmition = new ArrayList<String>();
							inputToken = new ArrayList<String>();
							currentOperationStage = 3;
							break;
						}
						case(3):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.size()==((availableNode.size()+1)*2))
								{
									MasterX = new Double(inputToken.get(0));
									MasterY = new Double(inputToken.get(1));
									for(int dummy = 1;dummy<(availableNode.size()+1);dummy++)
									{
										baseX.add(new Double(inputToken.get(dummy*2)));
										baseY.add(new Double(inputToken.get(dummy*2+1)));
									}
									currentOperationStage = 4;
									SerialPortInputTransmition = new ArrayList<String>();
								}else {
									systemOut("Error Input, check and enter again\n");
									currentOperationStage = 2;
								}
							}
							inputToken = new ArrayList<String>();
							break;
						}
						case(4):
						{
							systemOut("Input the x, y coordinate for the Target node:\n");
							SerialPortInputTransmition = new ArrayList<String>();
							inputToken = new ArrayList<String>();
							currentOperationStage = 5;
							break;
						}
						case(5):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.size()==2)
								{
									TargetX = new Double(inputToken.get(0));
									TargetY = new Double(inputToken.get(1));
									SerialPortInputTransmition = new ArrayList<String>();
									sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
									currentOperationStage = 6;
								}else {
									systemOut("Error Input, check and enter again\n");
									currentOperationStage = 4;
								}
							}
							inputToken = new ArrayList<String>();
							break;
						}
						case(6):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("T2B PKR:"))
								{
									int nodeAddress = Integer.parseInt(line.substring(line.indexOf("addr:")+5,line.indexOf(" T2B")));
									if(nodeAddress==2)
									{
										systemOut("Excluded node2 Data\n");
									}else {
										RSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
										AlphaRSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("Alp:")+4, line.indexOf(" FS"))));
									}
								}
								if(line.contains("End"))
								{
									for(List<Double> eachRssiList:RSSIList)
									{
										if(eachRssiList.isEmpty())
										{
											systemOut("Some Nodes Not responce, retransmit\n");
											RSSIList = new ArrayList<List<Double>>(availableNode.size());
											AlphaRSSIList = new ArrayList<List<Double>>();
											for(int index = 0;index<availableNode.size();index++)
											{
												RSSIList.add(new ArrayList<Double>());
												AlphaRSSIList.add(new ArrayList<Double>());
											}
											sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
											SerialPortInputTransmition = new ArrayList<String>();
											currentOperationStage = 6;
											inputToken = new ArrayList<String>();
											break;
										}
									}
									systemOut("Start calculation\n");
									
									//nusty
									
									//node info
									File outputxlsxFile = new File(basedir.getAbsolutePath()+fileNameExtension.format(new Date())+"Book.xlsx");
									XSSFWorkbook outputworkbook = new XSSFWorkbook();
									POIXMLProperties xmlProps = outputworkbook.getProperties();    
									POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
									coreProps.setCreator("Pequod");
									XSSFSheet rawDataSheet = outputworkbook.createSheet("RawData");
									int rawDataSheetCurrentRow = 0;
									XSSFRow nodeNAME = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeNAME.createCell(0, CellType.STRING).setCellValue("Node info");
									nodeNAME.createCell(1, CellType.STRING).setCellValue("Node0");
									nodeNAME.createCell(2, CellType.STRING).setCellValue("Node1");
									for(int dummy = 0;dummy<availableNode.size();dummy++)
									{
										nodeNAME.createCell(dummy+3, CellType.STRING).setCellValue("Node"+availableNode.get(dummy).toString());
									}
									rawDataSheetCurrentRow++;
									XSSFRow nodex = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodex.createCell(0, CellType.STRING).setCellValue("x");
									nodex.createCell(1, CellType.NUMERIC).setCellValue(MasterX);
									nodex.createCell(2, CellType.NUMERIC).setCellValue(TargetX);
									XSSFRow nodey = rawDataSheet.createRow(rawDataSheetCurrentRow+1);
									nodey.createCell(0, CellType.STRING).setCellValue("y");
									nodey.createCell(1, CellType.NUMERIC).setCellValue(MasterY);
									nodey.createCell(2, CellType.NUMERIC).setCellValue(TargetY);
									for(int dummy = 0;dummy<availableNode.size();dummy++)
									{
										nodex.createCell(dummy+3, CellType.NUMERIC).setCellValue(baseX.get(dummy));
										nodey.createCell(dummy+3, CellType.NUMERIC).setCellValue(baseY.get(dummy));
									}
									rawDataSheetCurrentRow+=2;
									
									//Alpha info
									XSSFRow nodeAnchorName = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeAnchorName.createCell(0, CellType.STRING).setCellValue("RSSIforAlpha");
									nodeAnchorName.createCell(1, CellType.STRING).setCellValue("distance to master");
									nodeAnchorName.createCell(2, CellType.STRING).setCellValue("Mean");
									nodeAnchorName.createCell(3, CellType.STRING).setCellValue("Values in each");
									rawDataSheetCurrentRow++;
									List<Double> AlphaRSSImean = new ArrayList<Double>();
									for(int index = 0;index<availableNode.size();index++)
									{
										double mean = PequodsMaths.getMean(PequodsMaths.removeOutliner(AlphaRSSIList.get(index)));
										AlphaRSSImean.add(mean);
										systemOut("RSSI of master to node address:"+availableNode.get(index)+" x:"+baseX.get(index)+" y:"+baseY.get(index)+" RSSI:"+mean+"\n");
										XSSFRow nodeAnchorAlphaRSSI = rawDataSheet.createRow(rawDataSheetCurrentRow);
										nodeAnchorAlphaRSSI.createCell(0, CellType.STRING).setCellValue("Node"+availableNode.get(index));
										nodeAnchorAlphaRSSI.createCell(1, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(MasterX-baseX.get(index).doubleValue(), 2)+Math.pow(MasterX-baseY.get(index).doubleValue(), 2)));
										nodeAnchorAlphaRSSI.createCell(2, CellType.NUMERIC).setCellValue(AlphaRSSImean.get(index));
										int CRow = 3;
										for(Double AlphaRSSI:AlphaRSSIList.get(index))
										{
											nodeAnchorAlphaRSSI.createCell(CRow, CellType.NUMERIC).setCellValue(AlphaRSSI.doubleValue());
											CRow++;
										}
										rawDataSheetCurrentRow++;
									}
									double[] AlphaN1m = PequodsMaths.calculateAlphaWithCustomData(AlphaRSSImean, MasterX, MasterY, baseX, baseY);
									alpha = AlphaN1m[0];
									RSSI1m = AlphaN1m[1];
									systemOut("R square of Alpha: "+AlphaN1m[2]+"\n");
									XSSFSheet summaryDataSheet = outputworkbook.createSheet("Summary");
									int summaryDataSheetCurrentRow = 0;
									XSSFRow alphaInfoName = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									alphaInfoName.createCell(0, CellType.STRING).setCellValue("1mRSSI");
									alphaInfoName.createCell(1, CellType.STRING).setCellValue("alpha");
									alphaInfoName.createCell(2, CellType.STRING).setCellValue("R square");
									summaryDataSheetCurrentRow++;
									XSSFRow alphaInfoData = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									alphaInfoData.createCell(0, CellType.NUMERIC).setCellValue(AlphaN1m[1]);
									alphaInfoData.createCell(1, CellType.NUMERIC).setCellValue(AlphaN1m[0]);
									alphaInfoData.createCell(2, CellType.NUMERIC).setCellValue(AlphaN1m[2]); 
									summaryDataSheetCurrentRow++;
									
									//nodeTargetRSSI
									XSSFRow nodeTargetToAnchorName = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeTargetToAnchorName.createCell(0, CellType.STRING).setCellValue("RSSIforTarget");
									nodeTargetToAnchorName.createCell(1, CellType.STRING).setCellValue("Distance to Target");
									nodeTargetToAnchorName.createCell(2, CellType.STRING).setCellValue("Ideal RSSI");
									nodeTargetToAnchorName.createCell(3, CellType.STRING).setCellValue("Mean");
									nodeTargetToAnchorName.createCell(4, CellType.STRING).setCellValue("Values in each");
									rawDataSheetCurrentRow++;
									List<Double> RSSImean = new ArrayList<Double>();
									for(int index = 0;index<availableNode.size();index++)
									{
										double mean = PequodsMaths.getMean(PequodsMaths.removeOutliner(RSSIList.get(index)));
										RSSImean.add(mean);
										systemOut("node address:"+availableNode.get(index)+" x:"+baseX.get(index)+" y:"+baseY.get(index)+" RSSI:"+mean+"\n");
										XSSFRow nodeAnchorTargetRSSI = rawDataSheet.createRow(rawDataSheetCurrentRow);
										nodeAnchorTargetRSSI.createCell(0, CellType.STRING).setCellValue("Node"+availableNode.get(index));
										nodeAnchorTargetRSSI.createCell(1, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-baseX.get(index).doubleValue(), 2)+Math.pow(TargetY-baseY.get(index).doubleValue(), 2)));
										nodeAnchorTargetRSSI.createCell(2, CellType.NUMERIC).setCellValue(AlphaN1m[0]-AlphaN1m[1]*10*Math.log10(Math.sqrt(Math.pow(TargetX-baseX.get(index).doubleValue(), 2)+Math.pow(TargetY-baseY.get(index).doubleValue(), 2))));
										nodeAnchorTargetRSSI.createCell(3, CellType.NUMERIC).setCellValue(RSSImean.get(index));
										int CRow = 4;
										for(Double AlphaRSSI:RSSIList.get(index))
										{
											nodeAnchorTargetRSSI.createCell(CRow, CellType.NUMERIC).setCellValue(AlphaRSSI.doubleValue());
											CRow++;
										}
										rawDataSheetCurrentRow++;
									}
									systemOut("1m RSSI: "+RSSI1m+"\n");
									systemOut("Alpha: "+alpha+"\n");
									double[] result = PequodsMaths.LLSForLocalization(RSSImean, baseX, baseY, RSSI1m, alpha);
									
									//Localisation
									List<List<Integer>> nodesList = new ArrayList<List<Integer>>();
									List<Double> locationxList = new ArrayList<Double>(); 
									List<Double> locationyList = new ArrayList<Double>(); 
									List<Double> locationDeltaxList = new ArrayList<Double>(); 
									List<Double> locationDeltayList = new ArrayList<Double>(); 
									List<ClusterableLocation> locations = new ArrayList<ClusterableLocation>();
									List<Double> nodeMBREList = new ArrayList<Double>(); 
									List<Double> nodeDeltaList = new ArrayList<Double>();
									List<Double> RsquareList = new ArrayList<Double>(); 
									List<Double> LEList = new ArrayList<Double>(); 
									XSSFRow resultNodesRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultNodesRow.createCell(0, CellType.STRING).setCellValue("Localisation Result using:");
									rawDataSheetCurrentRow++;
									XSSFRow resultXRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultXRow.createCell(0, CellType.STRING).setCellValue("x");
									rawDataSheetCurrentRow++;
									XSSFRow resultYRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultYRow.createCell(0, CellType.STRING).setCellValue("y");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaXRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaXRow.createCell(0, CellType.STRING).setCellValue("Delta x");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaYRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaYRow.createCell(0, CellType.STRING).setCellValue("Delta y");
									rawDataSheetCurrentRow++;
									XSSFRow resultMBRERow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultMBRERow.createCell(0, CellType.STRING).setCellValue("MBRE");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaRow.createCell(0, CellType.STRING).setCellValue("LLS Delta");
									rawDataSheetCurrentRow++;
									XSSFRow resultRsquareRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultRsquareRow.createCell(0, CellType.STRING).setCellValue("Rsquare");
									rawDataSheetCurrentRow++;
									XSSFRow resultLERow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultLERow.createCell(0, CellType.STRING).setCellValue("LE");
									rawDataSheetCurrentRow++;
									//all nodes
									systemOut("all nodes:\n");
									systemOut("x: "+result[0]+"\n");
									systemOut("y: "+result[1]+"\n");
									systemOut("x^2+y^2: "+result[2]+"\n");
									systemOut("R square: "+result[3]+"\n");
									systemOut("MBRE: "+result[4]+"\n");
									XSSFRow summaryInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									summaryInfoRow.createCell(0, CellType.STRING).setCellValue("Method");
									summaryInfoRow.createCell(1, CellType.STRING).setCellValue("x");
									summaryInfoRow.createCell(2, CellType.STRING).setCellValue("y");
									summaryInfoRow.createCell(3, CellType.STRING).setCellValue("Delta x");
									summaryInfoRow.createCell(4, CellType.STRING).setCellValue("Delta y");
									summaryInfoRow.createCell(5, CellType.STRING).setCellValue("LE");
									summaryDataSheetCurrentRow++;
									
									XSSFRow targetInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									targetInfoRow.createCell(0, CellType.STRING).setCellValue("Target");
									targetInfoRow.createCell(1, CellType.NUMERIC).setCellValue(TargetX);
									targetInfoRow.createCell(2, CellType.NUMERIC).setCellValue(TargetY);
									summaryDataSheetCurrentRow++;
									
									nodesList.add(availableNode);
									StringBuilder allNodeString = new StringBuilder(nodesList.get(0).get(0).toString());
									for(int dummy = 1;dummy<nodesList.get(0).size();dummy++)
									{
										allNodeString.append(","+nodesList.get(0).get(dummy).toString());
									}
									System.out.println(allNodeString.toString());
									resultNodesRow.createCell(nodesList.size(), CellType.STRING).setCellValue(allNodeString.toString());
									locationxList.add(new Double(result[0]));
									resultXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(result[0]);
									locationyList.add(new Double(result[1]));
									resultYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(result[1]);
									locationDeltaxList.add(new Double(result[0]-TargetX));
									resultDeltaXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(result[0]-TargetX);
									locationDeltayList.add(new Double(result[1]-TargetY));
									resultDeltaYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(result[1]-TargetY);
									locations.add(new ClusterableLocation(result[0], result[1]));
									nodeMBREList.add(result[4]);
									resultMBRERow.createCell(nodeMBREList.size(), CellType.NUMERIC).setCellValue(result[4]);
									nodeDeltaList.add(new Double(Math.sqrt(Math.abs(Math.pow(result[0], 2)+Math.pow(result[1], 2)-result[2]))));
									resultDeltaRow.createCell(nodeDeltaList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.abs(Math.pow(result[0], 2)+Math.pow(result[1], 2)-result[2])));
									RsquareList.add(result[3]);
									resultRsquareRow.createCell(RsquareList.size(), CellType.NUMERIC).setCellValue(result[3]);
									LEList.add(new Double(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2))));
									resultLERow.createCell(LEList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2)));
									XSSFRow allNodeRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									allNodeRow.createCell(0, CellType.STRING).setCellValue("all Nodes");
									allNodeRow.createCell(1, CellType.NUMERIC).setCellValue(result[0]);
									allNodeRow.createCell(2, CellType.NUMERIC).setCellValue(result[1]);
									allNodeRow.createCell(3, CellType.NUMERIC).setCellValue(result[0]-TargetX);
									allNodeRow.createCell(4, CellType.NUMERIC).setCellValue(result[1]-TargetY);
									allNodeRow.createCell(5, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2)));
									summaryDataSheetCurrentRow++;
									
									//other combination
									for(int noOfNode = 3;noOfNode<RSSImean.size();noOfNode++)
									{
										int noOfCombination = (int)CombinatoricsUtils.binomialCoefficient(RSSImean.size(), noOfNode);
										Iterator<int[]> listOfCom = CombinatoricsUtils.combinationsIterator(RSSImean.size(), noOfNode);
										int[][] Combination = new int[noOfCombination][noOfNode];
										int index = 0;
										
										while(listOfCom.hasNext())
										{
											List<Integer> nodes = new ArrayList<Integer>();
											Combination[index] = listOfCom.next();
											for(int no = 0;no<noOfNode;no++)
											{
												nodes.add(availableNode.get(Combination[index][no]));
												systemOut(availableNode.get(Combination[index][no])+",");
											}
											nodesList.add(nodes);
											StringBuilder NodeStringFor = new StringBuilder(nodes.get(0).toString());
											for(int dummy = 1;dummy<nodes.size();dummy++)
											{
												NodeStringFor.append(","+nodes.get(dummy).toString());
											}
											resultNodesRow.createCell(nodesList.size(), CellType.STRING).setCellValue(NodeStringFor.toString());
											systemOut("\n");
											
											List<Double> RSSImeanFor = new ArrayList<Double>();
											List<Double> baseXFor = new ArrayList<Double>();
											List<Double> baseYFor = new ArrayList<Double>();
											for(int nodeNo:Combination[index])
											{
												RSSImeanFor.add(RSSImean.get(nodeNo));
												baseXFor.add(baseX.get(nodeNo));
												baseYFor.add(baseY.get(nodeNo));
											}
											double[] resultFor = PequodsMaths.LLSForLocalization(RSSImeanFor, baseXFor, baseYFor, RSSI1m, alpha);
											systemOut("x: "+resultFor[0]+"\n");
											systemOut("y: "+resultFor[1]+"\n");
											systemOut("x^2+y^2: "+resultFor[2]+"\n");
											systemOut("R square: "+resultFor[3]+"\n");
											systemOut("MBRE: "+resultFor[4]+"\n");
											locationxList.add(new Double(resultFor[0]));
											resultXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(resultFor[0]);
											locationyList.add(new Double(resultFor[1]));
											resultYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(resultFor[1]);
											locationDeltaxList.add(new Double(resultFor[0]-TargetX));
											resultDeltaXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(resultFor[0]-TargetX);
											locationDeltayList.add(new Double(resultFor[1]-TargetY));
											resultDeltaYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(resultFor[1]-TargetY);
											locations.add(new ClusterableLocation(resultFor[0], resultFor[1]));
											nodeMBREList.add(resultFor[4]);
											resultMBRERow.createCell(nodeMBREList.size(), CellType.NUMERIC).setCellValue(resultFor[4]);
											nodeDeltaList.add(new Double(Math.sqrt(Math.abs(Math.pow(resultFor[0], 2)+Math.pow(resultFor[1], 2)-resultFor[2]))));
											resultDeltaRow.createCell(nodeDeltaList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.abs(Math.pow(resultFor[0], 2)+Math.pow(resultFor[1], 2)-resultFor[2])));
											RsquareList.add(resultFor[3]);
											resultRsquareRow.createCell(RsquareList.size(), CellType.NUMERIC).setCellValue(resultFor[3]);
											LEList.add(new Double(Math.sqrt(Math.pow(TargetX-resultFor[0], 2)+Math.pow(TargetY-resultFor[1], 2))));
											resultLERow.createCell(LEList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-resultFor[0], 2)+Math.pow(TargetY-resultFor[1], 2)));
											index++;
										}
										
									}
									
									
									//PLOT LE, MBRE, Delta, Rsquare
									XSSFSheet LEMBREDeltaSheet = outputworkbook.createSheet("LEMBREDeltaRsqeare");
									int LEMBREDeltaSheetCurrentRow = 16;
									int LEMBREDeltaSheetCurrentColumn = 0;
									PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "LE", LEList, locationxList, locationyList, TargetX, TargetY);
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									double[] MBREl = PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "MBRE", nodeMBREList, locationxList, locationyList, TargetX, TargetY);
									XSSFRow mMBREInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									mMBREInfoRow.createCell(0, CellType.STRING).setCellValue("m.MBRE");
									mMBREInfoRow.createCell(1, CellType.NUMERIC).setCellValue(MBREl[0]);
									mMBREInfoRow.createCell(2, CellType.NUMERIC).setCellValue(MBREl[1]);
									mMBREInfoRow.createCell(3, CellType.NUMERIC).setCellValue(MBREl[0]-TargetX);
									mMBREInfoRow.createCell(4, CellType.NUMERIC).setCellValue(MBREl[1]-TargetY);
									mMBREInfoRow.createCell(5, CellType.NUMERIC).setCellValue(MBREl[2]);
									summaryDataSheetCurrentRow++;
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									double[] Deltal = PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "Delta", nodeDeltaList, locationxList, locationyList, TargetX, TargetY);
									XSSFRow mDeltaInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									mDeltaInfoRow.createCell(0, CellType.STRING).setCellValue("m.LLS Delta");
									mDeltaInfoRow.createCell(1, CellType.NUMERIC).setCellValue(Deltal[0]);
									mDeltaInfoRow.createCell(2, CellType.NUMERIC).setCellValue(Deltal[1]);
									mDeltaInfoRow.createCell(3, CellType.NUMERIC).setCellValue(Deltal[0]-TargetX);
									mDeltaInfoRow.createCell(4, CellType.NUMERIC).setCellValue(Deltal[1]-TargetY);
									mDeltaInfoRow.createCell(5, CellType.NUMERIC).setCellValue(Deltal[2]);
									summaryDataSheetCurrentRow++;
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									
									
									//vMBRE vDelta vCombine
									List<Double> MBREvotes = PequodsMaths.normaliseVote(nodeMBREList, 25.0);
									double[] vMBREl = PequodsMaths.voting(MBREvotes, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vMBREInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vMBREInfoRow.createCell(0, CellType.STRING).setCellValue("MBRE vote");
									vMBREInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vMBREl[0]);
									vMBREInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vMBREl[1]);
									vMBREInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vMBREl[0]-TargetX);
									vMBREInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vMBREl[1]-TargetY);
									vMBREInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vMBREl[2]);
									summaryDataSheetCurrentRow++;
									List<Double> Deltavotes = PequodsMaths.normaliseVote(nodeDeltaList, 25.0);
									double[] vDeltal = PequodsMaths.voting(Deltavotes, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vDeltaInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDeltaInfoRow.createCell(0, CellType.STRING).setCellValue("Delta vote");
									vDeltaInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vDeltal[0]);
									vDeltaInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vDeltal[1]);
									vDeltaInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vDeltal[0]-TargetX);
									vDeltaInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vDeltal[1]-TargetY);
									vDeltaInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vDeltal[2]);
									summaryDataSheetCurrentRow++;
									List<Double> combineVote = new ArrayList<Double>();
									for(int dummy = 0;dummy<MBREvotes.size();dummy++)
									{
										combineVote.add(new Double(MBREvotes.get(dummy)/2.0+Deltavotes.get(dummy)/2.0));
									}
									double[] vCombinel = PequodsMaths.voting(combineVote, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vCombineInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vCombineInfoRow.createCell(0, CellType.STRING).setCellValue("Combine vote");
									vCombineInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vCombinel[0]);
									vCombineInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vCombinel[1]);
									vCombineInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vCombinel[0]-TargetX);
									vCombineInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vCombinel[1]-TargetY);
									vCombineInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vCombinel[2]);
									summaryDataSheetCurrentRow++;
									
									//KMC3
									XSSFSheet KMC3Sheet = outputworkbook.createSheet("KMC3");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster3 = new KMeansPlusPlusClusterer<ClusterableLocation>(3, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster3 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster3, locations.size()*20);
									double[] KMC3l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster3, nodeMBREList, KMC3Sheet, TargetX, TargetY, nodesList, "KMC3");
									XSSFRow BNKMC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC3InfoRow.createCell(0, CellType.STRING).setCellValue("3KMC BestNote");
									BNKMC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC3l[0]);
									BNKMC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC3l[1]);
									BNKMC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC3l[0]-TargetX);
									BNKMC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC3l[1]-TargetY);
									BNKMC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC3l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC3InfoRow.createCell(0, CellType.STRING).setCellValue("3KMC vote");
									vKMC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC3l[3]);
									vKMC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC3l[4]);
									vKMC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC3l[3]-TargetX);
									vKMC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC3l[4]-TargetY);
									vKMC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC3l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC4
									XSSFSheet KMC4Sheet = outputworkbook.createSheet("KMC4");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster4 = new KMeansPlusPlusClusterer<ClusterableLocation>(4, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster4 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster4, locations.size()*20);
									double[] KMC4l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster4, nodeMBREList, KMC4Sheet, TargetX, TargetY, nodesList, "KMC4");
									XSSFRow BNKMC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC4InfoRow.createCell(0, CellType.STRING).setCellValue("4KMC BestNode");
									BNKMC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC4l[0]);
									BNKMC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC4l[1]);
									BNKMC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC4l[0]-TargetX);
									BNKMC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC4l[1]-TargetY);
									BNKMC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC4l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC4InfoRow.createCell(0, CellType.STRING).setCellValue("4KMC vote");
									vKMC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC4l[3]);
									vKMC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC4l[4]);
									vKMC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC4l[3]-TargetX);
									vKMC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC4l[4]-TargetY);
									vKMC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC4l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC5
									XSSFSheet KMC5Sheet = outputworkbook.createSheet("KMC5");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster5 = new KMeansPlusPlusClusterer<ClusterableLocation>(5, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster5 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster5, locations.size()*20);
									double[] KMC5l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster5, nodeMBREList, KMC5Sheet, TargetX, TargetY, nodesList, "KMC5");
									XSSFRow BNKMC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC5InfoRow.createCell(0, CellType.STRING).setCellValue("5KMC BestNode");
									BNKMC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC5l[0]);
									BNKMC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC5l[1]);
									BNKMC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC5l[0]-TargetX);
									BNKMC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC5l[1]-TargetY);
									BNKMC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC5l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC5InfoRow.createCell(0, CellType.STRING).setCellValue("5KMC vote");
									vKMC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC5l[3]);
									vKMC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC5l[4]);
									vKMC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC5l[3]-TargetX);
									vKMC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC5l[4]-TargetY);
									vKMC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC5l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC6
									XSSFSheet KMC6Sheet = outputworkbook.createSheet("KMC6");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster6 = new KMeansPlusPlusClusterer<ClusterableLocation>(6, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster6 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster6, locations.size()*20);
									double[] KMC6l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster6, nodeMBREList, KMC6Sheet, TargetX, TargetY, nodesList, "KMC6");
									XSSFRow BNKMC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC6InfoRow.createCell(0, CellType.STRING).setCellValue("6KMC BestNode");
									BNKMC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC6l[0]);
									BNKMC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC6l[1]);
									BNKMC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC6l[0]-TargetX);
									BNKMC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC6l[1]-TargetY);
									BNKMC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC6l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC6InfoRow.createCell(0, CellType.STRING).setCellValue("6KMC vote");
									vKMC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC6l[3]);
									vKMC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC6l[4]);
									vKMC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC6l[3]-TargetX);
									vKMC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC6l[4]-TargetY);
									vKMC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC6l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC7
									XSSFSheet KMC7Sheet = outputworkbook.createSheet("KMC7");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster7 = new KMeansPlusPlusClusterer<ClusterableLocation>(7, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster7 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster7, locations.size()*20);
									double[] KMC7l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster7, nodeMBREList, KMC7Sheet, TargetX, TargetY, nodesList, "KMC7");
									XSSFRow BNKMC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC7InfoRow.createCell(0, CellType.STRING).setCellValue("7KMC BestNode");
									BNKMC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC7l[0]);
									BNKMC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC7l[1]);
									BNKMC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC7l[0]-TargetX);
									BNKMC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC7l[1]-TargetY);
									BNKMC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC7l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC7InfoRow.createCell(0, CellType.STRING).setCellValue("7KMC vote");
									vKMC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC7l[3]);
									vKMC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC7l[4]);
									vKMC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC7l[3]-TargetX);
									vKMC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC7l[4]-TargetY);
									vKMC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC7l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC3
									XSSFSheet DBC3Sheet = outputworkbook.createSheet("DBC3");
									DBSCANClusterer<ClusterableLocation> DBCluster3 = new DBSCANClusterer<ClusterableLocation>(3.0, 5);
									double[] DBC3l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster3, nodeMBREList, DBC3Sheet, TargetX, TargetY, nodesList, "DBC3");
									XSSFRow BNDBC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC3InfoRow.createCell(0, CellType.STRING).setCellValue("DBC3 BestNode");
									BNDBC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC3l[0]);
									BNDBC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC3l[1]);
									BNDBC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC3l[0]-TargetX);
									BNDBC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC3l[1]-TargetY);
									BNDBC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC3l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC3InfoRow.createCell(0, CellType.STRING).setCellValue("DBC3 vote");
									vDBC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC3l[3]);
									vDBC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC3l[4]);
									vDBC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC3l[3]-TargetX);
									vDBC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC3l[4]-TargetY);
									vDBC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC3l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC4
									XSSFSheet DBC4Sheet = outputworkbook.createSheet("DBC4");
									DBSCANClusterer<ClusterableLocation> DBCluster4 = new DBSCANClusterer<ClusterableLocation>(4.0, 5);
									double[] DBC4l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster4, nodeMBREList, DBC4Sheet, TargetX, TargetY, nodesList, "DBC4");
									XSSFRow BNDBC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC4InfoRow.createCell(0, CellType.STRING).setCellValue("DBC4 BestNode");
									BNDBC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC4l[0]);
									BNDBC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC4l[1]);
									BNDBC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC4l[0]-TargetX);
									BNDBC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC4l[1]-TargetY);
									BNDBC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC4l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC4InfoRow.createCell(0, CellType.STRING).setCellValue("DBC4 vote");
									vDBC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC4l[3]);
									vDBC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC4l[4]);
									vDBC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC4l[3]-TargetX);
									vDBC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC4l[4]-TargetY);
									vDBC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC4l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC5
									XSSFSheet DBC5Sheet = outputworkbook.createSheet("DBC5");
									DBSCANClusterer<ClusterableLocation> DBCluster5 = new DBSCANClusterer<ClusterableLocation>(5.0, 5);
									double[] DBC5l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster5, nodeMBREList, DBC5Sheet, TargetX, TargetY, nodesList, "DBC5");
									XSSFRow BNDBC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC5InfoRow.createCell(0, CellType.STRING).setCellValue("DBC5 BestNode");
									BNDBC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC5l[0]);
									BNDBC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC5l[1]);
									BNDBC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC5l[0]-TargetX);
									BNDBC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC5l[1]-TargetY);
									BNDBC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC5l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC5InfoRow.createCell(0, CellType.STRING).setCellValue("DBC5 vote");
									vDBC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC5l[3]);
									vDBC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC5l[4]);
									vDBC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC5l[3]-TargetX);
									vDBC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC5l[4]-TargetY);
									vDBC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC5l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC6
									XSSFSheet DBC6Sheet = outputworkbook.createSheet("DBC6");
									DBSCANClusterer<ClusterableLocation> DBCluster6 = new DBSCANClusterer<ClusterableLocation>(6.0, 5);
									double[] DBC6l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster6, nodeMBREList, DBC6Sheet, TargetX, TargetY, nodesList, "DBC6");
									XSSFRow BNDBC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC6InfoRow.createCell(0, CellType.STRING).setCellValue("DBC6 BestNode");
									BNDBC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC6l[0]);
									BNDBC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC6l[1]);
									BNDBC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC6l[0]-TargetX);
									BNDBC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC6l[1]-TargetY);
									BNDBC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC6l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC6InfoRow.createCell(0, CellType.STRING).setCellValue("DBC6 vote");
									vDBC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC6l[3]);
									vDBC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC6l[4]);
									vDBC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC6l[3]-TargetX);
									vDBC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC6l[4]-TargetY);
									vDBC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC6l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC7
									XSSFSheet DBC7Sheet = outputworkbook.createSheet("DBC7");
									DBSCANClusterer<ClusterableLocation> DBCluster7 = new DBSCANClusterer<ClusterableLocation>(7.0, 5);
									double[] DBC7l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster7, nodeMBREList, DBC7Sheet, TargetX, TargetY, nodesList, "DBC7");
									XSSFRow BNDBC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC7InfoRow.createCell(0, CellType.STRING).setCellValue("DBC7 BestNode");
									BNDBC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC7l[0]);
									BNDBC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC7l[1]);
									BNDBC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC7l[0]-TargetX);
									BNDBC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC7l[1]-TargetY);
									BNDBC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC7l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC7InfoRow.createCell(0, CellType.STRING).setCellValue("DBC7 vote");
									vDBC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC7l[3]);
									vDBC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC7l[4]);
									vDBC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC7l[3]-TargetX);
									vDBC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC7l[4]-TargetY);
									vDBC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC7l[5]);
									summaryDataSheetCurrentRow++;
									
									FileOutputStream writeFile = new FileOutputStream(outputxlsxFile);
									outputworkbook.write(writeFile);
									outputworkbook.close();
									writeFile.close();
									
									systemOut("Calculation finished!\n");
									
									currentOperationStage = 0;
									currentOperation = idle;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						default:
						{
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
					}
					break;
				}
				case(measureAlphaNlocateTargetMirror):
				{
					switch(currentOperationStage)
					{
						case(0):
						{
							sendToSerial("discover\n");
							numberOfSample = (short)Integer.parseInt(inputToken.get(1));
							systemOut("locate target with "+numberOfSample+" sample\n");
							availableNode = new ArrayList<Integer>();
							baseX = new ArrayList<Double>();
							baseY = new ArrayList<Double>();
							currentOperationStage = 1;
							break;
						}
						case(1):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("Available Nodes: "))
								{
									String[] nodeAddress = line.substring(line.indexOf(", ")+2,line.length()-2).split(", ");
									initNum = nodeAddress.length;
									for(int index = 0;index<nodeAddress.length;index++)
									{
										if(Integer.parseInt(nodeAddress[index])==2)
										{
											systemOut("Excluded node2\n");
										}else {
											availableNode.add(Integer.parseInt(nodeAddress[index]));
										}
									}
								}
								if(line.contains("Total number of "))
								{
									//!!!!!!!!!!!!!change is if no. of node>10
									if(availableNode.size()>3)
									{
										//
										for(int index = 0;index<initNum;index++)
										{
											availableNode.add(initNum+availableNode.get(index));
										}
										//
										
										RSSIList = new ArrayList<List<Double>>(availableNode.size());
										AlphaRSSIList = new ArrayList<List<Double>>();
										for(int index = 0;index<availableNode.size();index++)
										{
											RSSIList.add(new ArrayList<Double>());
											AlphaRSSIList.add(new ArrayList<Double>());
										}
										
										currentOperationStage = 2;
									}else{
										systemOut("Not enough available node! Return to idle\n");
										currentOperationStage = 0;
										currentOperation = idle;
									}
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						case(2):
						{
							StringBuilder stringOut = new StringBuilder("Master Node");
							for(int dummy = 0;dummy<availableNode.size();dummy++)
							{
								stringOut.append(","+availableNode.get(dummy).toString());
							}
							systemOut("Input the x, y coordinate for BaseStation address: "+stringOut+"\n");
							SerialPortInputTransmition = new ArrayList<String>();
							inputToken = new ArrayList<String>();
							currentOperationStage = 3;
							break;
						}
						case(3):
						{
							if(!inputToken.isEmpty())
							{
								systemOut("sIZE: "+(availableNode.size()+1)*2+"\n");
								if(inputToken.size()==((availableNode.size()+1)*2))
								{
									MasterX = new Double(inputToken.get(0));
									MasterY = new Double(inputToken.get(1));
									for(int dummy = 1;dummy<(availableNode.size()+1);dummy++)
									{
										baseX.add(new Double(inputToken.get(dummy*2)));
										baseY.add(new Double(inputToken.get(dummy*2+1)));
									}
									currentOperationStage = 4;
									SerialPortInputTransmition = new ArrayList<String>();
								}else {
									systemOut("Error Input, check and enter again\n");
									currentOperationStage = 2;
								}
							}
							inputToken = new ArrayList<String>();
							break;
						}
						case(4):
						{
							systemOut("Input the x, y coordinate for the Target node:\n");
							SerialPortInputTransmition = new ArrayList<String>();
							inputToken = new ArrayList<String>();
							currentOperationStage = 5;
							break;
						}
						case(5):
						{
							if(!inputToken.isEmpty())
							{
								if(inputToken.size()==2)
								{
									TargetX = new Double(inputToken.get(0));
									TargetY = new Double(inputToken.get(1));
									SerialPortInputTransmition = new ArrayList<String>();
									sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
									currentOperationStage = 6;
								}else {
									systemOut("Error Input, check and enter again\n");
									currentOperationStage = 4;
								}
							}
							inputToken = new ArrayList<String>();
							break;
						}
						case(6):
						{
							Iterator<String> SerialPortInputTransmitions = SerialPortInputTransmition.iterator();
							while(SerialPortInputTransmitions.hasNext())
							{
								String line = SerialPortInputTransmitions.next();
								if(line.contains("T2B PKR:"))
								{
									if(mirrorFlag)//mirror here
									{
										int nodeAddress = (initNum+Integer.parseInt(line.substring(line.indexOf("addr:")+5,line.indexOf(" T2B"))));
										if(nodeAddress==2)
										{
											systemOut("Excluded node2 Data\n");
										}else {
											RSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
											AlphaRSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("Alp:")+4, line.indexOf(" FS"))));
										}
									}else {
										int nodeAddress = (Integer.parseInt(line.substring(line.indexOf("addr:")+5,line.indexOf(" T2B"))));
										if(nodeAddress==2)
										{
											systemOut("Excluded node2 Data\n");
										}else {
											RSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("PKR:")+4, line.indexOf(" SNR"))));
											AlphaRSSIList.get(availableNode.indexOf(new Integer(nodeAddress))).add(new Double(line.substring(line.indexOf("Alp:")+4, line.indexOf(" FS"))));
										}
									}
								}
								if(line.contains("End"))
								{
									//mirror here
									if(mirrorFlag)
									{
										mirrorFlag = false;
									}else {
										mirrorFlag = true;
										currentOperationStage = 4;
										break;
									}
									
									//check for validation;
									for(List<Double> eachRssiList:RSSIList)
									{
										if(eachRssiList.isEmpty())
										{
											systemOut("Some Nodes Not responce, retransmit\n");
											RSSIList = new ArrayList<List<Double>>(availableNode.size());
											AlphaRSSIList = new ArrayList<List<Double>>();
											for(int index = 0;index<availableNode.size();index++)
											{
												RSSIList.add(new ArrayList<Double>());
												AlphaRSSIList.add(new ArrayList<Double>());
											}
											sendToSerial("targettobasetransmitionn "+numberOfSample+"\n");
											SerialPortInputTransmition = new ArrayList<String>();
											currentOperationStage = 6;
											inputToken = new ArrayList<String>();
											break;
										}
									}
									systemOut("Start calculation\n");
									
									//node info
									File outputxlsxFile = new File(basedir.getAbsolutePath()+fileNameExtension.format(new Date())+"Book.xlsx");
									XSSFWorkbook outputworkbook = new XSSFWorkbook();
									POIXMLProperties xmlProps = outputworkbook.getProperties();    
									POIXMLProperties.CoreProperties coreProps =  xmlProps.getCoreProperties();
									coreProps.setCreator("Pequod");
									XSSFSheet rawDataSheet = outputworkbook.createSheet("RawData");
									int rawDataSheetCurrentRow = 0;
									XSSFRow nodeNAME = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeNAME.createCell(0, CellType.STRING).setCellValue("Node info");
									nodeNAME.createCell(1, CellType.STRING).setCellValue("Node0");
									nodeNAME.createCell(2, CellType.STRING).setCellValue("Node1");
									for(int dummy = 0;dummy<availableNode.size();dummy++)
									{
										nodeNAME.createCell(dummy+3, CellType.STRING).setCellValue("Node"+availableNode.get(dummy).toString());
									}
									rawDataSheetCurrentRow++;
									XSSFRow nodex = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodex.createCell(0, CellType.STRING).setCellValue("x");
									nodex.createCell(1, CellType.NUMERIC).setCellValue(MasterX);
									nodex.createCell(2, CellType.NUMERIC).setCellValue(TargetX);
									XSSFRow nodey = rawDataSheet.createRow(rawDataSheetCurrentRow+1);
									nodey.createCell(0, CellType.STRING).setCellValue("y");
									nodey.createCell(1, CellType.NUMERIC).setCellValue(MasterY);
									nodey.createCell(2, CellType.NUMERIC).setCellValue(TargetY);
									for(int dummy = 0;dummy<availableNode.size();dummy++)
									{
										nodex.createCell(dummy+3, CellType.NUMERIC).setCellValue(baseX.get(dummy));
										nodey.createCell(dummy+3, CellType.NUMERIC).setCellValue(baseY.get(dummy));
									}
									rawDataSheetCurrentRow+=2;
									
									//Alpha info
									XSSFRow nodeAnchorName = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeAnchorName.createCell(0, CellType.STRING).setCellValue("RSSIforAlpha");
									nodeAnchorName.createCell(1, CellType.STRING).setCellValue("distance to master");
									nodeAnchorName.createCell(2, CellType.STRING).setCellValue("Mean");
									nodeAnchorName.createCell(3, CellType.STRING).setCellValue("Values in each");
									rawDataSheetCurrentRow++;
									List<Double> AlphaRSSImean = new ArrayList<Double>();
									for(int index = 0;index<availableNode.size();index++)
									{
										double mean = PequodsMaths.getMean(PequodsMaths.removeOutliner(AlphaRSSIList.get(index)));
										AlphaRSSImean.add(mean);
										systemOut("RSSI of master to node address:"+availableNode.get(index)+" x:"+baseX.get(index)+" y:"+baseY.get(index)+" RSSI:"+mean+"\n");
										XSSFRow nodeAnchorAlphaRSSI = rawDataSheet.createRow(rawDataSheetCurrentRow);
										nodeAnchorAlphaRSSI.createCell(0, CellType.STRING).setCellValue("Node"+availableNode.get(index));
										nodeAnchorAlphaRSSI.createCell(1, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(MasterX-baseX.get(index).doubleValue(), 2)+Math.pow(MasterX-baseY.get(index).doubleValue(), 2)));
										nodeAnchorAlphaRSSI.createCell(2, CellType.NUMERIC).setCellValue(AlphaRSSImean.get(index));
										int CRow = 3;
										for(Double AlphaRSSI:AlphaRSSIList.get(index))
										{
											nodeAnchorAlphaRSSI.createCell(CRow, CellType.NUMERIC).setCellValue(AlphaRSSI.doubleValue());
											CRow++;
										}
										rawDataSheetCurrentRow++;
									}
									double[] AlphaN1m = PequodsMaths.calculateAlphaWithCustomData(AlphaRSSImean, MasterX, MasterY, baseX, baseY);
									alpha = AlphaN1m[0];
									RSSI1m = AlphaN1m[1];
									systemOut("R square of Alpha: "+AlphaN1m[2]+"\n");
									XSSFSheet summaryDataSheet = outputworkbook.createSheet("Summary");
									int summaryDataSheetCurrentRow = 0;
									XSSFRow alphaInfoName = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									alphaInfoName.createCell(0, CellType.STRING).setCellValue("1mRSSI");
									alphaInfoName.createCell(1, CellType.STRING).setCellValue("alpha");
									alphaInfoName.createCell(2, CellType.STRING).setCellValue("R square");
									summaryDataSheetCurrentRow++;
									XSSFRow alphaInfoData = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									alphaInfoData.createCell(0, CellType.NUMERIC).setCellValue(AlphaN1m[1]);
									alphaInfoData.createCell(1, CellType.NUMERIC).setCellValue(AlphaN1m[0]);
									alphaInfoData.createCell(2, CellType.NUMERIC).setCellValue(AlphaN1m[2]); 
									summaryDataSheetCurrentRow++;
									
									//nodeTargetRSSI
									XSSFRow nodeTargetToAnchorName = rawDataSheet.createRow(rawDataSheetCurrentRow);
									nodeTargetToAnchorName.createCell(0, CellType.STRING).setCellValue("RSSIforTarget");
									nodeTargetToAnchorName.createCell(1, CellType.STRING).setCellValue("Distance to Target");
									nodeTargetToAnchorName.createCell(2, CellType.STRING).setCellValue("Ideal RSSI");
									nodeTargetToAnchorName.createCell(3, CellType.STRING).setCellValue("Mean");
									nodeTargetToAnchorName.createCell(4, CellType.STRING).setCellValue("Values in each");
									rawDataSheetCurrentRow++;
									List<Double> RSSImean = new ArrayList<Double>();
									for(int index = 0;index<availableNode.size();index++)
									{
										double mean = PequodsMaths.getMean(PequodsMaths.removeOutliner(RSSIList.get(index)));
										RSSImean.add(mean);
										systemOut("node address:"+availableNode.get(index)+" x:"+baseX.get(index)+" y:"+baseY.get(index)+" RSSI:"+mean+"\n");
										XSSFRow nodeAnchorTargetRSSI = rawDataSheet.createRow(rawDataSheetCurrentRow);
										nodeAnchorTargetRSSI.createCell(0, CellType.STRING).setCellValue("Node"+availableNode.get(index));
										nodeAnchorTargetRSSI.createCell(1, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-baseX.get(index).doubleValue(), 2)+Math.pow(TargetY-baseY.get(index).doubleValue(), 2)));
										nodeAnchorTargetRSSI.createCell(2, CellType.NUMERIC).setCellValue(AlphaN1m[0]-AlphaN1m[1]*10*Math.log10(Math.sqrt(Math.pow(TargetX-baseX.get(index).doubleValue(), 2)+Math.pow(TargetY-baseY.get(index).doubleValue(), 2))));
										nodeAnchorTargetRSSI.createCell(3, CellType.NUMERIC).setCellValue(RSSImean.get(index));
										int CRow = 4;
										for(Double AlphaRSSI:RSSIList.get(index))
										{
											nodeAnchorTargetRSSI.createCell(CRow, CellType.NUMERIC).setCellValue(AlphaRSSI.doubleValue());
											CRow++;
										}
										rawDataSheetCurrentRow++;
									}
									systemOut("1m RSSI: "+RSSI1m+"\n");
									systemOut("Alpha: "+alpha+"\n");
									double[] result = PequodsMaths.LLSForLocalization(RSSImean, baseX, baseY, RSSI1m, alpha);
									
									//Localisation
									List<List<Integer>> nodesList = new ArrayList<List<Integer>>();
									List<Double> locationxList = new ArrayList<Double>(); 
									List<Double> locationyList = new ArrayList<Double>(); 
									List<Double> locationDeltaxList = new ArrayList<Double>(); 
									List<Double> locationDeltayList = new ArrayList<Double>(); 
									List<ClusterableLocation> locations = new ArrayList<ClusterableLocation>();
									List<Double> nodeMBREList = new ArrayList<Double>(); 
									List<Double> nodeDeltaList = new ArrayList<Double>();
									List<Double> RsquareList = new ArrayList<Double>(); 
									List<Double> LEList = new ArrayList<Double>(); 
									XSSFRow resultNodesRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultNodesRow.createCell(0, CellType.STRING).setCellValue("Localisation Result using:");
									rawDataSheetCurrentRow++;
									XSSFRow resultXRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultXRow.createCell(0, CellType.STRING).setCellValue("x");
									rawDataSheetCurrentRow++;
									XSSFRow resultYRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultYRow.createCell(0, CellType.STRING).setCellValue("y");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaXRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaXRow.createCell(0, CellType.STRING).setCellValue("Delta x");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaYRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaYRow.createCell(0, CellType.STRING).setCellValue("Delta y");
									rawDataSheetCurrentRow++;
									XSSFRow resultMBRERow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultMBRERow.createCell(0, CellType.STRING).setCellValue("MBRE");
									rawDataSheetCurrentRow++;
									XSSFRow resultDeltaRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultDeltaRow.createCell(0, CellType.STRING).setCellValue("LLS sqrt Delta");
									rawDataSheetCurrentRow++;
									XSSFRow resultRsquareRow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultRsquareRow.createCell(0, CellType.STRING).setCellValue("Rsquare");
									rawDataSheetCurrentRow++;
									XSSFRow resultLERow = rawDataSheet.createRow(rawDataSheetCurrentRow);
									resultLERow.createCell(0, CellType.STRING).setCellValue("LE");
									rawDataSheetCurrentRow++;
									//all nodes
									systemOut("all nodes:\n");
									systemOut("x: "+result[0]+"\n");
									systemOut("y: "+result[1]+"\n");
									systemOut("x^2+y^2: "+result[2]+"\n");
									systemOut("R square: "+result[3]+"\n");
									systemOut("MBRE: "+result[4]+"\n");
									XSSFRow summaryInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									summaryInfoRow.createCell(0, CellType.STRING).setCellValue("Method");
									summaryInfoRow.createCell(1, CellType.STRING).setCellValue("x");
									summaryInfoRow.createCell(2, CellType.STRING).setCellValue("y");
									summaryInfoRow.createCell(3, CellType.STRING).setCellValue("Delta x");
									summaryInfoRow.createCell(4, CellType.STRING).setCellValue("Delta y");
									summaryInfoRow.createCell(5, CellType.STRING).setCellValue("LE");
									summaryDataSheetCurrentRow++;
									
									XSSFRow targetInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									targetInfoRow.createCell(0, CellType.STRING).setCellValue("Target");
									targetInfoRow.createCell(1, CellType.NUMERIC).setCellValue(TargetX);
									targetInfoRow.createCell(2, CellType.NUMERIC).setCellValue(TargetY);
									summaryDataSheetCurrentRow++;
									
									nodesList.add(availableNode);
									StringBuilder allNodeString = new StringBuilder(nodesList.get(0).get(0).toString());
									for(int dummy = 1;dummy<nodesList.get(0).size();dummy++)
									{
										allNodeString.append(","+nodesList.get(0).get(dummy).toString());
									}
									System.out.println(allNodeString.toString());
									resultNodesRow.createCell(nodesList.size(), CellType.STRING).setCellValue(allNodeString.toString());
									locationxList.add(new Double(result[0]));
									resultXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(result[0]);
									locationyList.add(new Double(result[1]));
									resultYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(result[1]);
									locationDeltaxList.add(new Double(result[0]-TargetX));
									resultDeltaXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(result[0]-TargetX);
									locationDeltayList.add(new Double(result[1]-TargetY));
									resultDeltaYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(result[1]-TargetY);
									locations.add(new ClusterableLocation(result[0], result[1]));
									nodeMBREList.add(result[4]);
									resultMBRERow.createCell(nodeMBREList.size(), CellType.NUMERIC).setCellValue(result[4]);
									nodeDeltaList.add(new Double(Math.abs(Math.pow(result[0], 2)+Math.pow(result[1], 2)-result[2])));
									resultDeltaRow.createCell(nodeDeltaList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.abs(Math.pow(result[0], 2)+Math.pow(result[1], 2)-result[2])));
									RsquareList.add(result[3]);
									resultRsquareRow.createCell(RsquareList.size(), CellType.NUMERIC).setCellValue(result[3]);
									LEList.add(new Double(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2))));
									resultLERow.createCell(LEList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2)));
									XSSFRow allNodeRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									allNodeRow.createCell(0, CellType.STRING).setCellValue("all Nodes");
									allNodeRow.createCell(1, CellType.NUMERIC).setCellValue(result[0]);
									allNodeRow.createCell(2, CellType.NUMERIC).setCellValue(result[1]);
									allNodeRow.createCell(3, CellType.NUMERIC).setCellValue(result[0]-TargetX);
									allNodeRow.createCell(4, CellType.NUMERIC).setCellValue(result[1]-TargetY);
									allNodeRow.createCell(5, CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-result[0], 2)+Math.pow(TargetY-result[1], 2)));
									summaryDataSheetCurrentRow++;
									
									//other combination
									for(int noOfNode = 3;noOfNode<RSSImean.size();noOfNode++)
									{
										int noOfCombination = (int)CombinatoricsUtils.binomialCoefficient(RSSImean.size(), noOfNode);
										Iterator<int[]> listOfCom = CombinatoricsUtils.combinationsIterator(RSSImean.size(), noOfNode);
										int[][] Combination = new int[noOfCombination][noOfNode];
										int index = 0;
										
										while(listOfCom.hasNext())
										{
											List<Integer> nodes = new ArrayList<Integer>();
											Combination[index] = listOfCom.next();
											for(int no = 0;no<noOfNode;no++)
											{
												nodes.add(availableNode.get(Combination[index][no]));
												systemOut(availableNode.get(Combination[index][no])+",");
											}
											nodesList.add(nodes);
											StringBuilder NodeStringFor = new StringBuilder(nodes.get(0).toString());
											for(int dummy = 1;dummy<nodes.size();dummy++)
											{
												NodeStringFor.append(","+nodes.get(dummy).toString());
											}
											resultNodesRow.createCell(nodesList.size(), CellType.STRING).setCellValue(NodeStringFor.toString());
											systemOut("\n");
											
											List<Double> RSSImeanFor = new ArrayList<Double>();
											List<Double> baseXFor = new ArrayList<Double>();
											List<Double> baseYFor = new ArrayList<Double>();
											for(int nodeNo:Combination[index])
											{
												RSSImeanFor.add(RSSImean.get(nodeNo));
												baseXFor.add(baseX.get(nodeNo));
												baseYFor.add(baseY.get(nodeNo));
											}
											double[] resultFor = PequodsMaths.LLSForLocalization(RSSImeanFor, baseXFor, baseYFor, RSSI1m, alpha);
											systemOut("x: "+resultFor[0]+"\n");
											systemOut("y: "+resultFor[1]+"\n");
											systemOut("x^2+y^2: "+resultFor[2]+"\n");
											systemOut("R square: "+resultFor[3]+"\n");
											systemOut("MBRE: "+resultFor[4]+"\n");
											locationxList.add(new Double(resultFor[0]));
											resultXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(resultFor[0]);
											locationyList.add(new Double(resultFor[1]));
											resultYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(resultFor[1]);
											locationDeltaxList.add(new Double(resultFor[0]-TargetX));
											resultDeltaXRow.createCell(locationxList.size(), CellType.NUMERIC).setCellValue(resultFor[0]-TargetX);
											locationDeltayList.add(new Double(resultFor[1]-TargetY));
											resultDeltaYRow.createCell(locationyList.size(), CellType.NUMERIC).setCellValue(resultFor[1]-TargetY);
											locations.add(new ClusterableLocation(resultFor[0], resultFor[1]));
											nodeMBREList.add(resultFor[4]);
											resultMBRERow.createCell(nodeMBREList.size(), CellType.NUMERIC).setCellValue(resultFor[4]);
											nodeDeltaList.add(new Double(Math.sqrt(Math.abs(Math.pow(resultFor[0], 2)+Math.pow(resultFor[1], 2)-resultFor[2]))));
											resultDeltaRow.createCell(nodeDeltaList.size(), CellType.NUMERIC).setCellValue(Math.abs(Math.pow(resultFor[0], 2)+Math.pow(resultFor[1], 2)-resultFor[2]));
											RsquareList.add(resultFor[3]);
											resultRsquareRow.createCell(RsquareList.size(), CellType.NUMERIC).setCellValue(resultFor[3]);
											LEList.add(new Double(Math.sqrt(Math.pow(TargetX-resultFor[0], 2)+Math.pow(TargetY-resultFor[1], 2))));
											resultLERow.createCell(LEList.size(), CellType.NUMERIC).setCellValue(Math.sqrt(Math.pow(TargetX-resultFor[0], 2)+Math.pow(TargetY-resultFor[1], 2)));
											index++;
										}
										
									}
									
									
									//PLOT LE, MBRE, Delta, Rsquare
									XSSFSheet LEMBREDeltaSheet = outputworkbook.createSheet("LEMBREDeltaRsqeare");
									int LEMBREDeltaSheetCurrentRow = 16;
									int LEMBREDeltaSheetCurrentColumn = 0;
									PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "LE", LEList, locationxList, locationyList, TargetX, TargetY);
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									double[] MBREl = PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "MBRE", nodeMBREList, locationxList, locationyList, TargetX, TargetY);
									XSSFRow mMBREInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									mMBREInfoRow.createCell(0, CellType.STRING).setCellValue("m.MBRE");
									mMBREInfoRow.createCell(1, CellType.NUMERIC).setCellValue(MBREl[0]);
									mMBREInfoRow.createCell(2, CellType.NUMERIC).setCellValue(MBREl[1]);
									mMBREInfoRow.createCell(3, CellType.NUMERIC).setCellValue(MBREl[0]-TargetX);
									mMBREInfoRow.createCell(4, CellType.NUMERIC).setCellValue(MBREl[1]-TargetY);
									mMBREInfoRow.createCell(5, CellType.NUMERIC).setCellValue(MBREl[2]);
									summaryDataSheetCurrentRow++;
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									double[] Deltal = PequodsMaths.fourQScatterPlot(LEMBREDeltaSheet, LEMBREDeltaSheetCurrentRow, LEMBREDeltaSheetCurrentColumn, "Delta", nodeDeltaList, locationxList, locationyList, TargetX, TargetY);
									XSSFRow mDeltaInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									mDeltaInfoRow.createCell(0, CellType.STRING).setCellValue("m.LLS Delta");
									mDeltaInfoRow.createCell(1, CellType.NUMERIC).setCellValue(Deltal[0]);
									mDeltaInfoRow.createCell(2, CellType.NUMERIC).setCellValue(Deltal[1]);
									mDeltaInfoRow.createCell(3, CellType.NUMERIC).setCellValue(Deltal[0]-TargetX);
									mDeltaInfoRow.createCell(4, CellType.NUMERIC).setCellValue(Deltal[1]-TargetY);
									mDeltaInfoRow.createCell(5, CellType.NUMERIC).setCellValue(Deltal[2]);
									summaryDataSheetCurrentRow++;
									LEMBREDeltaSheetCurrentRow+=10;
									LEMBREDeltaSheetCurrentColumn+=7;
									
									
									//vMBRE vDelta vCombine
									List<Double> MBREvotes = PequodsMaths.normaliseVote(nodeMBREList, 25.0);
									double[] vMBREl = PequodsMaths.voting(MBREvotes, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vMBREInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vMBREInfoRow.createCell(0, CellType.STRING).setCellValue("MBRE vote");
									vMBREInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vMBREl[0]);
									vMBREInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vMBREl[1]);
									vMBREInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vMBREl[0]-TargetX);
									vMBREInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vMBREl[1]-TargetY);
									vMBREInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vMBREl[2]);
									summaryDataSheetCurrentRow++;
									List<Double> Deltavotes = PequodsMaths.normaliseVote(nodeDeltaList, 25.0);
									double[] vDeltal = PequodsMaths.voting(Deltavotes, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vDeltaInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDeltaInfoRow.createCell(0, CellType.STRING).setCellValue("Delta vote");
									vDeltaInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vDeltal[0]);
									vDeltaInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vDeltal[1]);
									vDeltaInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vDeltal[0]-TargetX);
									vDeltaInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vDeltal[1]-TargetY);
									vDeltaInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vDeltal[2]);
									summaryDataSheetCurrentRow++;
									List<Double> combineVote = new ArrayList<Double>();
									for(int dummy = 0;dummy<MBREvotes.size();dummy++)
									{
										combineVote.add(new Double(MBREvotes.get(dummy)/2.0+Deltavotes.get(dummy)/2.0));
									}
									double[] vCombinel = PequodsMaths.voting(combineVote, locationxList, locationyList, TargetX, TargetY);
									XSSFRow vCombineInfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vCombineInfoRow.createCell(0, CellType.STRING).setCellValue("Combine vote");
									vCombineInfoRow.createCell(1, CellType.NUMERIC).setCellValue(vCombinel[0]);
									vCombineInfoRow.createCell(2, CellType.NUMERIC).setCellValue(vCombinel[1]);
									vCombineInfoRow.createCell(3, CellType.NUMERIC).setCellValue(vCombinel[0]-TargetX);
									vCombineInfoRow.createCell(4, CellType.NUMERIC).setCellValue(vCombinel[1]-TargetY);
									vCombineInfoRow.createCell(5, CellType.NUMERIC).setCellValue(vCombinel[2]);
									summaryDataSheetCurrentRow++;
									
									//KMC3
									XSSFSheet KMC3Sheet = outputworkbook.createSheet("KMC3");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster3 = new KMeansPlusPlusClusterer<ClusterableLocation>(3, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster3 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster3, locations.size()*20);
									double[] KMC3l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster3, nodeMBREList, KMC3Sheet, TargetX, TargetY, nodesList, "KMC3");
									XSSFRow BNKMC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC3InfoRow.createCell(0, CellType.STRING).setCellValue("3KMC BestNote");
									BNKMC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC3l[0]);
									BNKMC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC3l[1]);
									BNKMC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC3l[0]-TargetX);
									BNKMC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC3l[1]-TargetY);
									BNKMC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC3l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC3InfoRow.createCell(0, CellType.STRING).setCellValue("3KMC vote");
									vKMC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC3l[3]);
									vKMC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC3l[4]);
									vKMC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC3l[3]-TargetX);
									vKMC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC3l[4]-TargetY);
									vKMC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC3l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC4
									XSSFSheet KMC4Sheet = outputworkbook.createSheet("KMC4");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster4 = new KMeansPlusPlusClusterer<ClusterableLocation>(4, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster4 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster4, locations.size()*20);
									double[] KMC4l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster4, nodeMBREList, KMC4Sheet, TargetX, TargetY, nodesList, "KMC4");
									XSSFRow BNKMC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC4InfoRow.createCell(0, CellType.STRING).setCellValue("4KMC BestNode");
									BNKMC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC4l[0]);
									BNKMC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC4l[1]);
									BNKMC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC4l[0]-TargetX);
									BNKMC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC4l[1]-TargetY);
									BNKMC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC4l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC4InfoRow.createCell(0, CellType.STRING).setCellValue("4KMC vote");
									vKMC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC4l[3]);
									vKMC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC4l[4]);
									vKMC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC4l[3]-TargetX);
									vKMC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC4l[4]-TargetY);
									vKMC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC4l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC5
									XSSFSheet KMC5Sheet = outputworkbook.createSheet("KMC5");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster5 = new KMeansPlusPlusClusterer<ClusterableLocation>(5, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster5 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster5, locations.size()*20);
									double[] KMC5l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster5, nodeMBREList, KMC5Sheet, TargetX, TargetY, nodesList, "KMC5");
									XSSFRow BNKMC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC5InfoRow.createCell(0, CellType.STRING).setCellValue("5KMC BestNode");
									BNKMC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC5l[0]);
									BNKMC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC5l[1]);
									BNKMC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC5l[0]-TargetX);
									BNKMC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC5l[1]-TargetY);
									BNKMC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC5l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC5InfoRow.createCell(0, CellType.STRING).setCellValue("5KMC vote");
									vKMC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC5l[3]);
									vKMC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC5l[4]);
									vKMC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC5l[3]-TargetX);
									vKMC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC5l[4]-TargetY);
									vKMC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC5l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC6
									XSSFSheet KMC6Sheet = outputworkbook.createSheet("KMC6");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster6 = new KMeansPlusPlusClusterer<ClusterableLocation>(6, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster6 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster6, locations.size()*20);
									double[] KMC6l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster6, nodeMBREList, KMC6Sheet, TargetX, TargetY, nodesList, "KMC6");
									XSSFRow BNKMC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC6InfoRow.createCell(0, CellType.STRING).setCellValue("6KMC BestNode");
									BNKMC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC6l[0]);
									BNKMC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC6l[1]);
									BNKMC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC6l[0]-TargetX);
									BNKMC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC6l[1]-TargetY);
									BNKMC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC6l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC6InfoRow.createCell(0, CellType.STRING).setCellValue("6KMC vote");
									vKMC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC6l[3]);
									vKMC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC6l[4]);
									vKMC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC6l[3]-TargetX);
									vKMC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC6l[4]-TargetY);
									vKMC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC6l[5]);
									summaryDataSheetCurrentRow++;
									
									//KMC7
									XSSFSheet KMC7Sheet = outputworkbook.createSheet("KMC7");
									KMeansPlusPlusClusterer<ClusterableLocation> kmeanCluster7 = new KMeansPlusPlusClusterer<ClusterableLocation>(7, locations.size()*10);
									MultiKMeansPlusPlusClusterer<ClusterableLocation> MkmeanCluster7 = new MultiKMeansPlusPlusClusterer<ClusterableLocation>(kmeanCluster7, locations.size()*20);
									double[] KMC7l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, MkmeanCluster7, nodeMBREList, KMC7Sheet, TargetX, TargetY, nodesList, "KMC7");
									XSSFRow BNKMC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNKMC7InfoRow.createCell(0, CellType.STRING).setCellValue("7KMC BestNode");
									BNKMC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC7l[0]);
									BNKMC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC7l[1]);
									BNKMC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC7l[0]-TargetX);
									BNKMC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC7l[1]-TargetY);
									BNKMC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC7l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vKMC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vKMC7InfoRow.createCell(0, CellType.STRING).setCellValue("7KMC vote");
									vKMC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(KMC7l[3]);
									vKMC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(KMC7l[4]);
									vKMC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(KMC7l[3]-TargetX);
									vKMC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(KMC7l[4]-TargetY);
									vKMC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(KMC7l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC3
									XSSFSheet DBC3Sheet = outputworkbook.createSheet("DBC3");
									DBSCANClusterer<ClusterableLocation> DBCluster3 = new DBSCANClusterer<ClusterableLocation>(3.0, 5);
									double[] DBC3l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster3, nodeMBREList, DBC3Sheet, TargetX, TargetY, nodesList, "DBC3");
									XSSFRow BNDBC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC3InfoRow.createCell(0, CellType.STRING).setCellValue("DBC3 BestNode");
									BNDBC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC3l[0]);
									BNDBC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC3l[1]);
									BNDBC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC3l[0]-TargetX);
									BNDBC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC3l[1]-TargetY);
									BNDBC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC3l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC3InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC3InfoRow.createCell(0, CellType.STRING).setCellValue("DBC3 vote");
									vDBC3InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC3l[3]);
									vDBC3InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC3l[4]);
									vDBC3InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC3l[3]-TargetX);
									vDBC3InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC3l[4]-TargetY);
									vDBC3InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC3l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC4
									XSSFSheet DBC4Sheet = outputworkbook.createSheet("DBC4");
									DBSCANClusterer<ClusterableLocation> DBCluster4 = new DBSCANClusterer<ClusterableLocation>(4.0, 5);
									double[] DBC4l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster4, nodeMBREList, DBC4Sheet, TargetX, TargetY, nodesList, "DBC4");
									XSSFRow BNDBC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC4InfoRow.createCell(0, CellType.STRING).setCellValue("DBC4 BestNode");
									BNDBC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC4l[0]);
									BNDBC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC4l[1]);
									BNDBC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC4l[0]-TargetX);
									BNDBC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC4l[1]-TargetY);
									BNDBC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC4l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC4InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC4InfoRow.createCell(0, CellType.STRING).setCellValue("DBC4 vote");
									vDBC4InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC4l[3]);
									vDBC4InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC4l[4]);
									vDBC4InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC4l[3]-TargetX);
									vDBC4InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC4l[4]-TargetY);
									vDBC4InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC4l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC5
									XSSFSheet DBC5Sheet = outputworkbook.createSheet("DBC5");
									DBSCANClusterer<ClusterableLocation> DBCluster5 = new DBSCANClusterer<ClusterableLocation>(5.0, 5);
									double[] DBC5l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster5, nodeMBREList, DBC5Sheet, TargetX, TargetY, nodesList, "DBC5");
									XSSFRow BNDBC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC5InfoRow.createCell(0, CellType.STRING).setCellValue("DBC5 BestNode");
									BNDBC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC5l[0]);
									BNDBC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC5l[1]);
									BNDBC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC5l[0]-TargetX);
									BNDBC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC5l[1]-TargetY);
									BNDBC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC5l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC5InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC5InfoRow.createCell(0, CellType.STRING).setCellValue("DBC5 vote");
									vDBC5InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC5l[3]);
									vDBC5InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC5l[4]);
									vDBC5InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC5l[3]-TargetX);
									vDBC5InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC5l[4]-TargetY);
									vDBC5InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC5l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC6
									XSSFSheet DBC6Sheet = outputworkbook.createSheet("DBC6");
									DBSCANClusterer<ClusterableLocation> DBCluster6 = new DBSCANClusterer<ClusterableLocation>(6.0, 5);
									double[] DBC6l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster6, nodeMBREList, DBC6Sheet, TargetX, TargetY, nodesList, "DBC6");
									XSSFRow BNDBC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC6InfoRow.createCell(0, CellType.STRING).setCellValue("DBC6 BestNode");
									BNDBC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC6l[0]);
									BNDBC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC6l[1]);
									BNDBC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC6l[0]-TargetX);
									BNDBC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC6l[1]-TargetY);
									BNDBC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC6l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC6InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC6InfoRow.createCell(0, CellType.STRING).setCellValue("DBC6 vote");
									vDBC6InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC6l[3]);
									vDBC6InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC6l[4]);
									vDBC6InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC6l[3]-TargetX);
									vDBC6InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC6l[4]-TargetY);
									vDBC6InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC6l[5]);
									summaryDataSheetCurrentRow++;
									
									//DBC7
									XSSFSheet DBC7Sheet = outputworkbook.createSheet("DBC7");
									DBSCANClusterer<ClusterableLocation> DBCluster7 = new DBSCANClusterer<ClusterableLocation>(7.0, 5);
									double[] DBC7l = PequodsMaths.bestClusterBestNodeNVotePlot(locations, DBCluster7, nodeMBREList, DBC7Sheet, TargetX, TargetY, nodesList, "DBC7");
									XSSFRow BNDBC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									BNDBC7InfoRow.createCell(0, CellType.STRING).setCellValue("DBC7 BestNode");
									BNDBC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC7l[0]);
									BNDBC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC7l[1]);
									BNDBC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC7l[0]-TargetX);
									BNDBC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC7l[1]-TargetY);
									BNDBC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC7l[2]);
									summaryDataSheetCurrentRow++;
									XSSFRow vDBC7InfoRow = summaryDataSheet.createRow(summaryDataSheetCurrentRow);
									vDBC7InfoRow.createCell(0, CellType.STRING).setCellValue("DBC7 vote");
									vDBC7InfoRow.createCell(1, CellType.NUMERIC).setCellValue(DBC7l[3]);
									vDBC7InfoRow.createCell(2, CellType.NUMERIC).setCellValue(DBC7l[4]);
									vDBC7InfoRow.createCell(3, CellType.NUMERIC).setCellValue(DBC7l[3]-TargetX);
									vDBC7InfoRow.createCell(4, CellType.NUMERIC).setCellValue(DBC7l[4]-TargetY);
									vDBC7InfoRow.createCell(5, CellType.NUMERIC).setCellValue(DBC7l[5]);
									summaryDataSheetCurrentRow++;
									
									FileOutputStream writeFile = new FileOutputStream(outputxlsxFile);
									outputworkbook.write(writeFile);
									outputworkbook.close();
									writeFile.close();
									
									systemOut("Calculation finished!\n");
									
									currentOperationStage = 0;
									currentOperation = idle;
								}
							}
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
						default:
						{
							inputToken = new ArrayList<String>();
							SerialPortInputTransmition = new ArrayList<String>();
							break;
						}
					}
				}
				default:
				{
					inputToken = new ArrayList<String>();
					SerialPortInputTransmition = new ArrayList<String>();
					break;
				}
			}
		}
		
		writeStringBuilderToFile(content, textLog);
		serialPortInputStream.close();
		serialPortOutputStream.close();
		serialPort.close();
		f.dispatchEvent(new WindowEvent(f, WindowEvent.WINDOW_CLOSING));
		f.dispose();
		return toExit;
	}
	private class SerialReader implements Runnable 
    {
		Scanner localScanner;
        
        public SerialReader (Scanner scanner)
        {
            this.localScanner = scanner;
        }
        
        public void run()
        {
        	while(true)
        	{
        		if(localScanner.hasNext())
        		{
        			input = new String(localScanner.nextLine());
        			newInput = true;
					if(input.equalsIgnoreCase("exit")||input.equalsIgnoreCase("end"))
					{
						break;
					}
				}
        	}
        }
    }
	public void sendToSerial(String commandToSend)
	{
		int commandLength = commandToSend.length();
		for(int i = 0;i<commandLength;i++)
		{
			try 
			{
				serialPortOutputStream.write((int)commandToSend.charAt(i));
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	public void systemOut(String messageToUser)
	{
		System.out.print(messageToUser);
		content.append(">>>>"+messageToUser);
	}
	public static void writeStringBuilderToFile(StringBuilder content, File file) throws IOException
	{
		FileWriter out = null;
		try
		{
			out=new FileWriter(file);
			file.createNewFile();
			out.write(content.toString());
			out.close();
		}finally 
		{
			if (out != null) 
			{
				out.close();
			}
		}
	}
}
