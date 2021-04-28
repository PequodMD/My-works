package localizationLoRa;

import org.apache.commons.math3.ml.clustering.Clusterable;

public class ClusterableLocation implements Clusterable 
{
	private double[] points;
	public ClusterableLocation(double x, double y)
	{
		points = new double[] {x, y};
	}
	@Override
	public double[] getPoint() 
	{
		return points;
	}

}
