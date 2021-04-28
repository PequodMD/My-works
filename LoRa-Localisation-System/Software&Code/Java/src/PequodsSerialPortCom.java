package localizationLoRa;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import gnu.io.CommPort;
import gnu.io.CommPortIdentifier;
import gnu.io.SerialPort;

public class PequodsSerialPortCom
{
	public static List<Object> getInOutputStream(Scanner scanner)
	{
		List<Object> inOutputStream = new ArrayList<Object>();
		Scanner inputScanner = scanner;
		List<String> listOfPortNames = getListOfPortNames();
		System.out.println("List of Available Serial Port:");
		Iterator<String> listOfPortNamesiterator = listOfPortNames.iterator();
		while(listOfPortNamesiterator.hasNext())
		{
			System.out.println(listOfPortNamesiterator.next());
		}
		System.out.println("Serial Port to be selected: ");
		String portName = inputScanner.nextLine();
		if(portName.equalsIgnoreCase("exit"))
		{
			inOutputStream.add(portName);
		}else{
			listOfPortNamesiterator = listOfPortNames.iterator();
			while(listOfPortNamesiterator.hasNext())
			{
				if(listOfPortNamesiterator.next().equals(portName))
				{
					try
		    		{
						inOutputStream.addAll(connect(portName));
		    		}catch ( Exception e ){
		    			e.printStackTrace();
		    		}
					break;
				}
			}
			if(inOutputStream.size()!=3)
			{
				System.out.println("No Such Serial Port, rescan 5 second later");
	    		for(int second = 0;second<5;second++)
	    		{
	    			System.out.print(".");
	    			try 
	    			{
	    				TimeUnit.SECONDS.sleep(1);
	    			} catch (InterruptedException e){
	    				e.printStackTrace();
	    			}
	    		}
	    		System.out.println();
			}
		}
		return inOutputStream;
	}
	public static List<String> getListOfPortNames()
    {
    	List<String> listOfPortNames = new ArrayList<String>();
        @SuppressWarnings("unchecked")
		java.util.Enumeration<CommPortIdentifier> portEnum = CommPortIdentifier.getPortIdentifiers();
        while ( portEnum.hasMoreElements() ) 
        {
            CommPortIdentifier portIdentifier = portEnum.nextElement();
            if(getPortTypeName(portIdentifier.getPortType()).equals("Serial")&&(!portIdentifier.isCurrentlyOwned()))
            {
            	listOfPortNames.add(portIdentifier.getName());
            }
        }   
        return listOfPortNames;
    }
	static String getPortTypeName ( int portType )
    {
        switch ( portType )
        {
            case CommPortIdentifier.PORT_I2C:
                return "I2C";
            case CommPortIdentifier.PORT_PARALLEL:
                return "Parallel";
            case CommPortIdentifier.PORT_RAW:
                return "Raw";
            case CommPortIdentifier.PORT_RS485:
                return "RS485";
            case CommPortIdentifier.PORT_SERIAL:
                return "Serial";
            default:
                return "unknown type";
        }
    }
	public static List<Object> connect(String portName)throws Exception
    {
		List<Object> inOutputStream = new ArrayList<Object>();
		CommPortIdentifier portIdentifier = CommPortIdentifier.getPortIdentifier(portName);
        if ( portIdentifier.isCurrentlyOwned())
        {
            System.out.println("Error: Port is currently in use");
        }else{
        	CommPort commPort;
        	try
        	{
        		commPort = portIdentifier.open(new String("Pequods"),2000);
        		if (commPort instanceof SerialPort)
                {
                	System.out.println("Connection Started, Enter END to end connection");
                    SerialPort serialPort = (SerialPort)commPort;
                    serialPort.setSerialPortParams(9600,SerialPort.DATABITS_8,SerialPort.STOPBITS_1,SerialPort.PARITY_NONE);
                    inOutputStream.add(serialPort.getInputStream());
                    inOutputStream.add(serialPort.getOutputStream());
                    inOutputStream.add(serialPort);
                }else{
                    System.out.println("Error: Only serial ports are handled.");
                }
        	}catch(gnu.io.PortInUseException e){
        		System.out.println("Error: Port is currently in use, press Enter key to rescan");
        	}
        }  
		return inOutputStream;
    }
}
