package localizationLoRa;

import java.io.IOException;
import java.util.List;
import java.util.Scanner;


public class Main 
{
	public static void main(String[] args) throws IOException 
	{
		Scanner inputScanner = new Scanner(System.in);
		while(true)
		{
			List<Object> inOutputStream = PequodsSerialPortCom.getInOutputStream(inputScanner);
			if(inOutputStream.size()==1)
			{
				if(((String)inOutputStream.get(0)).equalsIgnoreCase("exit"))
				{
					break;
				}
			}else if(inOutputStream.size()==3){
				if(new PequodsLoRaTerminal(inOutputStream).start(inputScanner))
				{
					break;
				}
			}
		}
		inputScanner.close();
	}
}
