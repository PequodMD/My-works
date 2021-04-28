package localizationLoRa;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.math3.ml.clustering.Cluster;
import org.apache.commons.math3.ml.clustering.Clusterer;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.commons.math3.stat.regression.OLSMultipleLinearRegression;
import org.apache.commons.math3.stat.regression.SimpleRegression;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFNoFillProperties;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFScatterChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class PequodsMaths 
{
	public static double calculateAlpha(double RSSI0, double RSSI1)
	{
		return (RSSI0-RSSI1)/10;
	}
	public static double[] calculateAlphaWithCustomData(List<Double> AlphaRSSImean, double MasterX, double MasterY, List<Double> baseX, List<Double> baseY)
	{
		double[] result = new double[3];
		SimpleRegression lineRegression = new SimpleRegression();
		double[][] data = new double[AlphaRSSImean.size()][];
		for(int index = 0;index<AlphaRSSImean.size();index++)
		{
			data[index] = new double[]{Math.log10(Math.sqrt(Math.pow(MasterX-baseX.get(index), 2)+Math.pow(MasterY-baseY.get(index), 2))), AlphaRSSImean.get(index)};
		}
		lineRegression.addData(data);
		double m = lineRegression.getSlope();
		result[0] = -m/10;//(b0-(log(10)*m+b0))/10
		result[1] = lineRegression.getIntercept();
		result[2] = lineRegression.getRSquare();
		return result;
	}
	public static double getMean(List<? extends Number> data)
	{
		DescriptiveStatistics dataStat = new DescriptiveStatistics(listToArray(data));
		return dataStat.getMean();
	}
	public static double[] LinealityTest(List<Double> nodeDistance, List<Double> nodeRSSI)
	{
		double[] result = new double[2];
		SimpleRegression lineRegression = new SimpleRegression();
		double[][] data = new double[nodeRSSI.size()][];
		for(int index = 0;index<nodeRSSI.size();index++)
		{
			data[index] = new double[]{Math.log10(nodeDistance.get(index)), nodeRSSI.get(index)};
		}
		lineRegression.addData(data);
		result[0] = lineRegression.getSlope();
		result[1] = lineRegression.getRSquare();
		return result;
	}
	public static double[] LLSForLocalization(List<Double> nodeRSSI, List<Double> nodeXcoordinate, List<Double> nodeYcoordinate, double Z0, double alpha)
	{
		double[] result = new double[5];//x, y, x2+y2, R2, MBRE
		OLSMultipleLinearRegression regression = new OLSMultipleLinearRegression();
		regression.setNoIntercept(true);
		
		double[] b = new double[4];
		double[][] A = new double[4][3];
		
		if(nodeRSSI.size()==3)
		{
			for(int index = 0;index<nodeRSSI.size();index++)
			{
				b[index] = Math.pow(10, (Z0-nodeRSSI.get(index).doubleValue())/(5.0*alpha))-(nodeXcoordinate.get(index).doubleValue()*nodeXcoordinate.get(index).doubleValue())-(nodeYcoordinate.get(index).doubleValue()*nodeYcoordinate.get(index).doubleValue());
				A[index][0] = nodeXcoordinate.get(index).doubleValue()*-2.0;
				A[index][1] = nodeYcoordinate.get(index).doubleValue()*-2.0;
				A[index][2] = 1.0;
			}
			b[3] = 0.0;
			A[3][0] = 0.0;
			A[3][1] = 0.0;
			A[3][2] = 0.0;
			regression.newSampleData(b, A);
		}else{
			b = new double[nodeRSSI.size()];
			A = new double[nodeRSSI.size()][3];
			for(int index = 0;index<nodeRSSI.size();index++)
			{
				b[index] = Math.pow(10, (Z0-nodeRSSI.get(index).doubleValue())/(5.0*alpha))-(nodeXcoordinate.get(index).doubleValue()*nodeXcoordinate.get(index).doubleValue())-(nodeYcoordinate.get(index).doubleValue()*nodeYcoordinate.get(index).doubleValue());
				A[index][0] = nodeXcoordinate.get(index).doubleValue()*-2.0;
				A[index][1] = nodeYcoordinate.get(index).doubleValue()*-2.0;
				A[index][2] = 1.0;
			}
			regression.newSampleData(b, A);
		}
		
		try
		{
			double[] theta = regression.estimateRegressionParameters();
			result[0] = theta[0];
			result[1] = theta[1];
			result[2] = theta[2];
			result[3] = regression.calculateRSquared();
		}catch (org.apache.commons.math3.linear.SingularMatrixException e){
			regression = new OLSMultipleLinearRegression();
			regression.setNoIntercept(true);
			if(nodeRSSI.size()==3)
			{
				for(int index = 0;index<nodeRSSI.size();index++)
				{
					double nX = nodeXcoordinate.get(index).doubleValue()+0.00000001*index;
					double nY = nodeYcoordinate.get(index).doubleValue()+0.00000001*index;;
					b[index] = Math.pow(10, (Z0-nodeRSSI.get(index).doubleValue())/(5.0*alpha))-(nX*nX)-(nY*nY);
					A[index][0] = nX*-2.0;
					A[index][1] = nY*-2.0;
					A[index][2] = 1.0;
				}
				b[3] = 0.0;
				A[3][0] = 0.0;
				A[3][1] = 0.0;
				A[3][2] = 0.0;
				regression.newSampleData(b, A);
			}else{
				b = new double[nodeRSSI.size()];
				A = new double[nodeRSSI.size()][3];
				for(int index = 0;index<nodeRSSI.size();index++)
				{
					double nX = nodeXcoordinate.get(index).doubleValue()+0.00000001*index;
					double nY = nodeYcoordinate.get(index).doubleValue()+0.00000001*index;
					b[index] = Math.pow(10, (Z0-nodeRSSI.get(index).doubleValue())/(5.0*alpha))-(nX*nX)-(nY*nY);
					A[index][0] = nX*-2.0;
					A[index][1] = nY*-2.0;
					A[index][2] = 1.0;
				}
				regression.newSampleData(b, A);
			}
			double[] theta = regression.estimateRegressionParameters();
			result[0] = theta[0];
			result[1] = theta[1];
			result[2] = theta[2];
			result[3] = regression.calculateRSquared();
		}
		double BRESum = 0.0;
		for(int node = 0;node<nodeRSSI.size();node++)
		{
			BRESum+=Math.abs(nodeRSSI.get(node)-Z0+(5.0*alpha*Math.log10(Math.pow(nodeXcoordinate.get(node)-result[0], 2)+Math.pow(nodeYcoordinate.get(node)-result[1], 2))));
		}
		result[4] = BRESum/nodeRSSI.size();
		
		return result;
	}
	public static List<Double> removeOutliner(List<Double> inputData)
	{
		List<Double> outputData = new ArrayList<Double>();
		DescriptiveStatistics dataStat = new DescriptiveStatistics(listToArray(inputData));
		for(int dummy = 0;dummy<inputData.size();dummy++)
		{
			if(Math.abs(inputData.get(dummy)-dataStat.getMean())<=2*dataStat.getStandardDeviation())
			{
				outputData.add(inputData.get(dummy));
			}
		}
		return outputData;
	}
	public static double[] listToArray(List<? extends Number> inputData)
	{
		double[] array = new double[inputData.size()];
		for(int dummy = 0;dummy<inputData.size();dummy++)
		{
			array[dummy] = inputData.get(dummy).doubleValue();
		}
		return array;
	}
	public static void scatterChat(XSSFSheet sheet, int startRow, int startColumn, String title, String x1Axis, String x2Axis, List<List<Double>> x1Lists, List<List<Double>> x2Lists, List<String> dataLabelList)
	{
		XSSFDrawing drawing = sheet.createDrawingPatriarch();
		XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startColumn, 0, startColumn+7, 16);
		XSSFChart chart = drawing.createChart(anchor);
		XDDFChartLegend legend = chart.getOrAddLegend();
		legend.setPosition(LegendPosition.TOP_RIGHT);
		XDDFValueAxis bottomAxis = chart.createValueAxis(AxisPosition.BOTTOM);
		bottomAxis.setTitle(x1Axis);
		bottomAxis.setMaximum(70);
		bottomAxis.setMinimum(-130);
		XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
		leftAxis.setTitle(x2Axis);
		leftAxis.setMaximum(70);
		leftAxis.setMinimum(-130);
		leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
		XDDFScatterChartData data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
		for(int dataPattenNo = 0;dataPattenNo<x1Lists.size();dataPattenNo++)
		{
			XSSFRow x1Row = sheet.createRow(startRow+dataPattenNo*2);
			XSSFRow x2Row = sheet.createRow(startRow+dataPattenNo*2+1);
			List<Double> x1List = x1Lists.get(dataPattenNo);
			List<Double> x2List = x2Lists.get(dataPattenNo);
			for(int dataNo = 0;dataNo<x1List.size();dataNo++)
			{
				XSSFCell x1Cell = x1Row.createCell(dataNo);
				x1Cell.setCellType(CellType.NUMERIC);
				x1Cell.setCellValue(x1List.get(dataNo));
				XSSFCell x2Cell = x2Row.createCell(dataNo);
				x2Cell.setCellType(CellType.NUMERIC);
				x2Cell.setCellValue(x2List.get(dataNo));
			}
			if(x1List.size()==0)
			{
				continue;
			}
			XDDFDataSource<Double> x1 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(startRow+dataPattenNo*2,startRow+dataPattenNo*2,0,x1List.size()-1));
			XDDFNumericalDataSource<Double> x2 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(startRow+dataPattenNo*2+1,startRow+dataPattenNo*2+1,0,x2List.size()-1));
			XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(x1, x2);
			series.setTitle(dataLabelList.get(dataPattenNo), null);
			series.setMarkerStyle(MarkerStyle.CIRCLE);
			series.setMarkerSize((short)5);
			XDDFNoFillProperties nofill = new XDDFNoFillProperties();
			XDDFLineProperties linepr = new XDDFLineProperties();
			linepr.setFillProperties(nofill);
			XDDFShapeProperties sharppr = series.getShapeProperties();
			if(sharppr==null)
			{
				sharppr = new XDDFShapeProperties();
			}
			sharppr.setLineProperties(linepr);
			series.setShapeProperties(sharppr);
		}
		chart.setTitleText(title);
		chart.plot(data);
	}
	public static double[] fourQScatterPlot(XSSFSheet sheet, int startRow, int startColumn, String title, List<Double> data, List<Double> xData, List<Double> yData, double targetx, double targety)
	{
		double[] minimun = new double[4];
		double[] dataArray = new double[data.size()];
		for(int dummy = 0;dummy<data.size();dummy++)
		{
			dataArray[dummy] = data.get(dummy);
		}
		DescriptiveStatistics LEStat = new DescriptiveStatistics(dataArray);
		List<Double> LEQ1xList = new ArrayList<Double>();
		List<Double> LEQ2xList = new ArrayList<Double>();
		List<Double> LEQ3xList = new ArrayList<Double>();
		List<Double> LEQ4xList = new ArrayList<Double>();
		List<Double> LEQ1yList = new ArrayList<Double>();
		List<Double> LEQ2yList = new ArrayList<Double>();
		List<Double> LEQ3yList = new ArrayList<Double>();
		List<Double> LEQ4yList = new ArrayList<Double>();
		for(int dummy = 0;dummy<data.size();dummy++)
		{
			if(data.get(dummy)<LEStat.getPercentile(25))
			{
				LEQ1xList.add(xData.get(dummy));
				LEQ1yList.add(yData.get(dummy));
				if(data.get(dummy)==LEStat.getMin())
				{
					minimun[0] = xData.get(dummy);
					minimun[1] = yData.get(dummy);
					minimun[2] = Math.sqrt(Math.pow(xData.get(dummy).doubleValue()-targetx, 2)+Math.pow(yData.get(dummy).doubleValue()-targety, 2));
				}
			}else if(data.get(dummy)>=LEStat.getPercentile(25)&&data.get(dummy)<LEStat.getPercentile(50)){
				LEQ2xList.add(xData.get(dummy));
				LEQ2yList.add(yData.get(dummy));
			}else if(data.get(dummy)>=LEStat.getPercentile(50)&&data.get(dummy)<LEStat.getPercentile(75)){
				LEQ3xList.add(xData.get(dummy));
				LEQ3yList.add(yData.get(dummy));
			}else {
				LEQ4xList.add(xData.get(dummy));
				LEQ4yList.add(yData.get(dummy));
			}
		}
		List<List<Double>> xLists = new ArrayList<List<Double>>();
		List<List<Double>> yLists = new ArrayList<List<Double>>();
		List<String> dataLableList = new ArrayList<String>();
		List<Double> tx = new ArrayList<Double>();
		tx.add(new Double(targetx));
		List<Double> ty = new ArrayList<Double>();
		ty.add(new Double(targety));

		dataLableList.add("Q4");
		xLists.add(LEQ4xList);
		yLists.add(LEQ4yList);
		dataLableList.add("Q3");
		xLists.add(LEQ3xList);
		yLists.add(LEQ3yList);
		dataLableList.add("Q2");
		xLists.add(LEQ2xList);
		yLists.add(LEQ2yList);
		dataLableList.add("Q1");
		xLists.add(LEQ1xList);
		yLists.add(LEQ1yList);
		dataLableList.add("Target");
		xLists.add(tx);
		yLists.add(ty);
		scatterChat(sheet, startRow, startColumn, title, "x", "y", xLists, yLists, dataLableList);
		return minimun;
	}
	public static List<Double> normaliseVote(List<Double> votes, double maxPercentile)
	{
		List<Double> result = new ArrayList<Double>();
		double[] voteArray = new double[votes.size()];
		for(int dummy = 0;dummy<votes.size();dummy++)
		{
			voteArray[dummy] = votes.get(dummy);
		}
		DescriptiveStatistics voteStat = new DescriptiveStatistics(voteArray);
		for(Double vote:votes)
		{
			if(vote<=voteStat.getPercentile(maxPercentile))
			{
				result.add(new Double(1.0-(vote-voteStat.getMin())/(voteStat.getPercentile(maxPercentile)-voteStat.getMin())));
			}else{
				result.add(new Double(0.0));
			}
		}
		return result;
	}
	public static double[] voting(List<Double> votes, List<Double> xData, List<Double> yData, double targetx, double targety)
	{
		double[] result = new double[3];
		result[0] = 0.0;
		result[1] = 0.0;
		double scoreSum = 0.0;
		for(int dummy = 0;dummy<votes.size();dummy++)
		{
			result[0] += votes.get(dummy)*xData.get(dummy);
			result[1] += votes.get(dummy)*yData.get(dummy);
			scoreSum += votes.get(dummy);
		}
		result[0] /= scoreSum;
		result[1] /= scoreSum;
		result[2] = Math.sqrt(Math.pow(result[0]-targetx, 2)+Math.pow(result[1]-targety, 2));
		return result;
	}
	public static double[] bestClusterBestNodeNVotePlot(List<ClusterableLocation> locations, Clusterer<ClusterableLocation> clusterer, List<Double> scoreDataList, XSSFSheet sheet, double targetx, double targety, List<List<Integer>> nodeLists, String plotTitle)
	{
		double[] results = new double[6];
		List<? extends Cluster<ClusterableLocation>> clusterResults = clusterer.cluster(locations);
		List<Double> clusterScoreList = new ArrayList<Double>();
		for(Cluster<ClusterableLocation> clusterResult:clusterResults)
		{
			Double Score = new Double(0.0);
			for(ClusterableLocation location:clusterResult.getPoints())
			{
				//System.out.println(location.getPoint()[0]+" ,"+location.getPoint()[1]);
				Score = new Double(scoreDataList.get(locations.indexOf(location)).doubleValue()+Score.doubleValue());
			}
			Score = new Double(Score.doubleValue()/(double)clusterResult.getPoints().size());
			clusterScoreList.add(Score);
		}
		double[] ScoreArray = new double[clusterScoreList.size()];
		for(int dummy = 0;dummy<clusterScoreList.size();dummy++)
		{
			ScoreArray[dummy] = clusterScoreList.get(dummy);
		}
		DescriptiveStatistics voteStat = new DescriptiveStatistics(ScoreArray);
		List<List<Double>> xLists = new ArrayList<List<Double>>();
		List<List<Double>> yLists = new ArrayList<List<Double>>();
		List<Double> bestClusterxList = new ArrayList<Double>();
		List<Double> bestClusteryList = new ArrayList<Double>();
		List<Double> otherClusterxList = new ArrayList<Double>();
		List<Double> otherClusteryList = new ArrayList<Double>();
		List<Double> targetxList = new ArrayList<Double>();
		List<Double> targetyList = new ArrayList<Double>();
		targetxList.add(new Double(targetx));
		targetyList.add(new Double(targety));
		List<Double> bestClusterScoreList = new ArrayList<Double>();
		for(int dummy = 0;dummy<clusterScoreList.size();dummy++)
		{
			if(clusterScoreList.get(dummy)==voteStat.getMin())
			{
				Cluster<ClusterableLocation> bestCluster = clusterResults.get(dummy);
				List<Integer> nodeVoteList = new ArrayList<Integer>(8);
				for(int inde = 0;inde<8;inde++)
				{
					nodeVoteList.add(new Integer(0));
				}
				for(ClusterableLocation location:bestCluster.getPoints())
				{
					bestClusterxList.add(location.getPoint()[0]);
					bestClusteryList.add(location.getPoint()[1]);
					bestClusterScoreList.add(scoreDataList.get(locations.indexOf(location)));
					for(Integer node:nodeLists.get(locations.indexOf(location)))
					{
						nodeVoteList.set(node.intValue()-3, new Integer(nodeVoteList.get(node.intValue()-3).intValue()+1));
					}
				}
				/*
				int dummy0 = 3;
				for(Integer nodeVote:nodeVoteList)
				{
					System.out.println("node: "+dummy0+" count:"+nodeVote.toString());
					dummy0++;
				}
				*/
				double[] nodeVoteArray = new double[nodeVoteList.size()];
				for(int index = 0;index<nodeVoteList.size();index++)
				{
					nodeVoteArray[index] = nodeVoteList.get(index);
				}
				List<Integer> goodNodeList = new ArrayList<Integer>();
				for(int index = 0;index<nodeVoteList.size();index++)
				{
					if(nodeVoteList.get(index) >= (new DescriptiveStatistics(nodeVoteArray).getMax()*0.8)||nodeVoteList.get(index) >= (new DescriptiveStatistics(nodeVoteArray).getPercentile(100*(1.0-3.0/(double)nodeVoteList.size()))))
					{
						goodNodeList.add(new Integer(index+3));
						//System.out.print(","+(index+3));
					}
				}
				//System.out.println();
				for(int index = 0;index<nodeLists.size();index++)
				{
					if(nodeLists.get(index).equals(goodNodeList))
					{
						//System.out.println("Hit!");
						results[0] = locations.get(index).getPoint()[0];
						results[1] = locations.get(index).getPoint()[1];
						results[2] = Math.sqrt(Math.pow(results[0]-targetx, 2)+Math.pow(results[1]-targety, 2));
						break;
					}
				}
				List<Double> normalisedScore = normaliseVote(bestClusterScoreList, 75.0);
				double[] votingl = voting(normalisedScore, bestClusterxList, bestClusteryList, targetx, targety);
				results[3]=votingl[0];
				results[4]=votingl[1];
				results[5]=votingl[2];
			}else{
				Cluster<ClusterableLocation> otherCluster = clusterResults.get(dummy);
				for(ClusterableLocation location:otherCluster.getPoints())
				{
					otherClusterxList.add(location.getPoint()[0]);
					otherClusteryList.add(location.getPoint()[1]);
				}
			}
		}
		xLists.add(otherClusterxList);
		yLists.add(otherClusteryList);
		xLists.add(bestClusterxList);
		yLists.add(bestClusteryList);
		xLists.add(targetxList);
		yLists.add(targetyList);
		List<String> dataLableList = new ArrayList<String>();
		dataLableList.add("Other Cluster");
		dataLableList.add("Best Cluster");
		dataLableList.add("Target");
		scatterChat(sheet, 0, 0, plotTitle, "x", "y", xLists, yLists, dataLableList);
		List<Integer> nodeVoteList = new ArrayList<Integer>(8);
		for(int inde = 0;inde<8;inde++)
		{
			nodeVoteList.add(new Integer(0));
		}
		return results;
	}
}
