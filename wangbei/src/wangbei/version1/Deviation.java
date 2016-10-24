package wangbei.version1;

public class Deviation {
	/**
	 * 传统的利用平均数求方差的方法,需要遍历数组两次
	 * @param a 目标数组
	 * @return 方差
	 */
	public static String ComputeVariance(Double a[]){
		double variance=0;//方差
		double average=0;//平均数
		int i,len=a.length;
		double sum=0,sum2=0;
		for(i=0;i<len;i++){
			sum+=a[i];
		}
		average=sum/len;
		for(i=0;i<len;i++){
			sum2+=(a[i]-average)*(a[i]-average);
		}
		variance=Math.pow(sum2/len, 0.5);
		return average +"," + variance;
	}
	
	public static String getMean(Double a[]){
		double average=0;//平均数
		int i,len=a.length;
		if(a.length == 0)
			return null;
		double sum=0;
		for(i=0;i<len;i++){
			if(a[i] == null)
				break;
			sum+=a[i];
		}
		return Double.toString(sum/len);
	}
	/**
	 * 只遍历数组一次求方差，利用公式DX^2=EX^2-(EX)^2
	 * @param doubles
	 * @return
	 */
	public static String ComputeVariance2(Double[] doubles){
		double variance=0;//方差
		double sum=0,sum2=0;
		int i=0,len=doubles.length;
		if(len == 0){
			return null;
		}
		for(;i<len;i++){
			if(doubles[i] == null)
				break;
			sum+=doubles[i];
			sum2+=doubles[i]*doubles[i];
		}
		variance=sum2/len-(sum/len)*(sum/len);
		return Double.toString(variance);
	}
	
}
