package wangbei.version1;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
/**
 * 该类是用于解析上传的excel文件，excel文件中的内容是基本是固定的，将解析后的数据进行初步处理后写入对应的文件中
 * 由于写入excel文件中单元格数量有限制，因为输出的格式以CSV结尾，这样就可以克服内存的限制
 * @author lvpenglin
 *
 */
public class ReadExcelToDB {
	/**
	 * 程序的入口
	 * @param args
	 * @throws Exception
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception, Exception {
		File fileC = new File(
				"C://Users//lvpenglin//Desktop//Computed_VARIAN-TRILOGY-1_6MV_Open00.xls");
		File fileM = new File(
				"C://Users//lvpenglin//Desktop//Measured_VARIAN-TRILOGY-1_6MV_Open.xls");
		Map mapC = readExcel(fileC);
		Map mapM = readExcel(fileM);
		Map mapMX = readExcelToXValue(fileM);
		Map mapCX = readExcelToXValue(fileC);
		computedInfo(mapC, mapM, mapCX, mapMX);
	}

	/**
	 * @param file 需要解析的excel文件路径
	 * @return  返回解析文件中的Y轴的值，每一个数据段封装在一个map中
	 * @throws BiffException
	 * @throws IOException
	 */
	public static Map readExcel(File file) throws BiffException, IOException {
		Workbook wb = Workbook.getWorkbook(file);// 从文件流中取得Excel工作区对象
		// 开始遍历Excel工作区
		Map map = new LinkedHashMap<String, List>();
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			// 依次获取工作区中的excel表格
			Sheet sheet = wb.getSheet(sheetIndex);
			double[] f = null;
			StringBuilder sb = new StringBuilder();
			// 遍历excel表格中的每一行，自定义从第二行开始，表格中第一行为表头
			int flag = 0;
			for (int i = 0; i < sheet.getRows(); i++) {
				String str0 = sheet.getCell(0, i).getContents();
				String temp;
				if (str0.contains("Fieldsize") || str0.contains("CurveType")
						|| str0.contains("StartPoint")) {
					if (str0.contains("Fieldsize")) {
						sb.append((++flag) + str0);
					} else {
						sb.append(str0);
					}
				}
				int j = 0;
				if (str0.contains("StartPoint")) {
					List list = new ArrayList<Double>();
					while (i++ < sheet.getRows()) {
						if (sheet.getCell(0, i).getContents()
								.equalsIgnoreCase("end")) {
							map.put(sb.toString(), list);
							sb = new StringBuilder();
							f = null;
							j = 0;
							break;
						} else {
							temp = sheet.getCell(0, i).getContents();
							double f_temp = Double.parseDouble(temp
									.substring(temp.indexOf(" ") + 1));// 截取从空格开始的后面所有的字符
							list.add(f_temp);
						}
					}
				}
				if (i == sheet.getRows() - 1) {// 判断是否定位到表格的最后一行
					break;
				}
			}
		}
		System.out.println(map.size());
		return map;
	}

	/**
	 * @param file  需要解析的excel文件路径
	 * @return  返回解析文件中的X轴的值，每一个数据段封装在一个map中
	 * @throws BiffException
	 * @throws IOException
	 */
	public static Map readExcelToXValue(File file) throws BiffException,
			IOException {
		Workbook wb = Workbook.getWorkbook(file);// 从文件流中取得Excel工作区对象
		// 开始遍历Excel工作区
		Map mapX = new LinkedHashMap<String, List>();
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			// 依次获取工作区中的excel表格
			Sheet sheet = wb.getSheet(sheetIndex);
			double[] f = null;
			StringBuilder sb = new StringBuilder();
			// 遍历excel表格中的每一行，自定义从第二行开始，表格中第一行为表头
			int flag = 0;
			for (int i = 0; i < sheet.getRows(); i++) {
				String str0 = sheet.getCell(0, i).getContents();
				String temp;
				if (str0.contains("Fieldsize") || str0.contains("CurveType")
						|| str0.contains("StartPoint")) {
					if (str0.contains("Fieldsize")) {
						sb.append((++flag) + str0);
					} else {
						sb.append(str0);
					}
				}
				int j = 0;
				if (str0.contains("StartPoint")) {
					List listX = new ArrayList<Double>();
					while (i++ < sheet.getRows()) {
						if (sheet.getCell(0, i).getContents()
								.equalsIgnoreCase("end")) {
							mapX.put(sb.toString(), listX);
							sb = new StringBuilder();
							f = null;
							j = 0;
							break;
						} else {
							temp = sheet.getCell(0, i).getContents();
							double f_tempX = Double.parseDouble(temp.substring(
									0, temp.indexOf(";")));// 截取从第一个字符开始到分号之前的所有的字符
							listX.add(f_tempX);
						}
					}
				}
				if (i == sheet.getRows() - 1) {// 判断是否定位到表格的最后一行
					break;
				}
			}
		}
		System.out.println(mapX.size());
		return mapX;
	}

	/**
	 * @param mapC  C情况下对应的Y的值
	 * @param mapM  M情况下对应的Y的值
	 * @param mapCX C情况下对应的X的值
	 * @param mapMX M情况下对应的X的值
	 * 核心的代码，对数据进行处理
	 * @throws Exception
	 */
	public static void computedInfo(Map mapC, Map mapM, Map mapCX, Map mapMX)
			throws Exception {
		Iterator<Map.Entry<String, List>> itC = mapC.entrySet().iterator();
		Iterator<Map.Entry<String, List>> itM = mapM.entrySet().iterator();
		Iterator<Map.Entry<String, List>> itCX = mapCX.entrySet().iterator();
		Iterator<Map.Entry<String, List>> itMX = mapMX.entrySet().iterator();
		Map resultMap = new LinkedHashMap<String, List>();
		List<Double> listLastLeft = new ArrayList<Double>();
		List<Double> listLastRight = new ArrayList<Double>();
		List<Double> listLastAB = new ArrayList<Double>();
		List<Double> listLastBC = new ArrayList<Double>();
		List<Double> listLastCD = new ArrayList<Double>();
		List<Double> listLastDE = new ArrayList<Double>();
		List<Double> listLastEF = new ArrayList<Double>();
		FileWriter fwprofile = new FileWriter(new File(
				"C://Users//lvpenglin//Desktop//结果profile.csv"));
		FileWriter fwpdd = new FileWriter(new File(
				"C://Users//lvpenglin//Desktop//结果pdd.csv"));
		int j = 0;
		while (itC.hasNext() && itM.hasNext() && itCX.hasNext()
				&& itMX.hasNext()) {
			Map.Entry<String, List> mapEleC = itC.next();
			Map.Entry<String, List> mapEleM = itM.next();
			Map.Entry<String, List> mapEleCX = itCX.next();
			Map.Entry<String, List> mapEleMX = itMX.next();

			String cKey = mapEleC.getKey();
			String mKey = mapEleM.getKey();

			String sCombine = cKey + mKey;

			List ListC = mapEleC.getValue();
			List ListM = mapEleM.getValue();
			List ListCX = mapEleCX.getValue();
			List ListMX = mapEleMX.getValue();
			if (sCombine.contains("Depth")) {
				double maxM = Collections.max(ListM);
				int index = ListM.indexOf(maxM);
				for (int i = 0; i < ListC.size() && i < ListM.size(); i++) {
					if (i <= ListM.indexOf(maxM)) {
						listLastLeft
								.add(((Double) ListC.get(i) - (Double) ListM
										.get(i)) / (Double) ListM.get(i));
					} else {
						listLastRight
								.add(((Double) ListC.get(i) - (Double) ListM
										.get(i)) / (Double) ListM.get(i));
					}
				}
				listLastRight.add(((Double) ListC.get(index) - (Double) ListM
						.get(index)) / (Double) ListM.get(index));
				fwpdd.write(sCombine
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastLeft
								.toArray(new Double[listLastLeft.size()]))
						+","
						+ Deviation.ComputeVariance((Double[]) listLastRight
								.toArray(new Double[listLastRight.size()]))
						+ "\r\n");
				fwpdd.flush();
				sCombine = new String();
				listLastLeft.clear();
				listLastRight.clear();
			} else {
				double midValue = 0;
				double CmidValue = 0;
				int mid = (int) Math.floor(ListM.size() * 0.5);
				if (ListM.size() % 2 == 1) {
					midValue = (double) ListM.get(mid);
				} else {
					midValue = ((double) ListM.get(mid) + (double) ListM
							.get(mid - 1)) / 2;
				}

				int midC = (int) Math.floor(ListC.size() * 0.5);
				if (ListC.size() % 2 == 1) {
					CmidValue = (double) ListC.get(mid);
				} else {
					CmidValue = ((double) ListC.get(mid) + (double) ListC
							.get(mid - 1)) / 2;
				}
				double A = (double) ListC.get(0);

				double B = (midValue * 0.2);

				double C = (midValue * 0.8);

				double D = (midValue * 0.8);

				double E = (midValue * 0.2);

				double F = (double) ListM.get(ListM.size()-1);

				double MA_indexValue = (midValue * 0.5);
				double MB_indexValue = (midValue * 0.9);
				double MC_indexValue = (midValue * 0.9);
				double MD_indexValue = (midValue * 0.5);
				int MA_index_X = 0;
				int MB_index_X = 0;
				int MC_index_X = 0;
				int MD_index_X = 0;
				double CA_indexValue = (CmidValue * 0.5);
				double CB_indexValue = (CmidValue * 0.9);
				double CC_indexValue = (CmidValue * 0.9);
				double CD_indexValue = (CmidValue * 0.5);
				int CA_index_X = 0;
				int CB_index_X = 0;
				int CC_index_X = 0;
				int CD_index_X = 0;

				for (int i = 0; i < ListC.size() && i < ListM.size(); i++) {
					if (i <= mid) {
						if ((Double) ListM.get(i) < MA_indexValue
								&& (Double) ListM.get(i + 1) >= MA_indexValue) {
							MA_index_X = i + 1;// A点在M中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListM.get(i) < MB_indexValue
								&& (Double) ListM.get(i + 1) >= MB_indexValue) {
							MB_index_X = i + 1;// B点在M中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListC.get(i) < CA_indexValue
								&& (Double) ListC.get(i + 1) >= CA_indexValue) {
							CA_index_X = i + 1;// A点在C中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListC.get(i) < CB_indexValue
								&& (Double) ListC.get(i + 1) >= CB_indexValue) {
							CB_index_X = i + 1;// B点在C中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListM.get(i) < B) {
							listLastAB
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / midValue);
							if ((Double) ListM.get(i + 1) >= B) {
								listLastAB
										.add(((Double) ListC.get(i + 1) - (Double) ListM
												.get(i + 1))
												/ midValue);
								continue;
							}
						} else if ((Double) ListM.get(i) < C) {
							listLastBC
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / (Double) ListM.get(i));
							if ((Double) ListM.get(i + 1) >= C) {
								listLastBC
										.add(((Double) ListC.get(i + 1) - (Double) ListM
												.get(i + 1))
												/ (Double) ListM.get(i + 1));
								continue;
							}
						} else {
							listLastCD
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / (Double) ListM.get(i));
						}
					} else {
						if ((Double) ListM.get(i) > MC_indexValue
								&& (Double) ListM.get(i + 1) <= MC_indexValue) {
							MC_index_X = i ;// C点在M中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListM.get(i) > MD_indexValue
								&& (Double) ListM.get(i + 1) <= MD_indexValue) {
							MD_index_X = i ;// D点在M中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListC.get(i) > CC_indexValue
								&& (Double) ListC.get(i + 1) <= CC_indexValue) {
							CC_index_X = i ;// C点在C中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListC.get(i) > CD_indexValue
								&& (Double) ListC.get(i + 1) <= CD_indexValue) {
							CD_index_X = i ;// D点在C中的index，需要求出对应的x坐标，x坐标的值应该单独存储在mapX中
						}
						if ((Double) ListM.get(i) >= D) {
							listLastCD
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / (Double) ListM.get(i));
							if ((Double) ListM.get(i) >= D
									&& (Double) ListM.get(i + 1) <= D) {
								listLastDE
										.add(((Double) ListC.get(i) - (Double) ListM
												.get(i))
												/ (Double) ListM.get(i));
							}
						} else if ((Double) ListM.get(i) >= E
								&& (Double) ListM.get(i) <= D) {
							listLastDE
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / (Double) ListM.get(i));
							if ((Double) ListM.get(i) >= E
									&& (Double) ListM.get(i + 1) <= E) {
								listLastEF
										.add(((Double) ListC.get(i) - (Double) ListM
												.get(i))
												/ midValue);
							}
						} else {
							listLastEF
									.add(((Double) ListC.get(i) - (Double) ListM
											.get(i)) / midValue);
						}
					}
				}
				fwprofile.write(sCombine
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastAB
								.toArray(new Double[listLastAB.size()]))
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastBC
								.toArray(new Double[listLastBC.size()]))
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastCD
								.toArray(new Double[listLastCD.size()]))
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastDE
								.toArray(new Double[listLastDE.size()]))
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastEF
								.toArray(new Double[listLastEF.size()]))
						+ ","
						+ ((Double) ListCX.get(CB_index_X) - (Double) ListCX
								.get(CA_index_X))
						+ ","
						+ ((Double) ListMX.get(MB_index_X) - (Double) ListMX
								.get(MA_index_X))
						+ ","
						+ (((Double) ListCX.get(CB_index_X) - (Double) ListCX
								.get(CA_index_X)) - ((Double) ListMX
								.get(MB_index_X) - (Double) ListMX
								.get(MA_index_X)))
						+ ","
						+ ((Double) ListCX.get(CD_index_X) - (Double) ListCX
								.get(CA_index_X))
								+ ","
						+ ((Double) ListMX.get(MD_index_X) - (Double) ListMX
								.get(MA_index_X))
						+ ","
						+ (((Double) ListCX.get(CD_index_X) - (Double) ListCX
								.get(CA_index_X)) - ((Double) ListMX
								.get(MD_index_X) - (Double) ListMX
								.get(MA_index_X)))
						+ ","
						+ ((Double) ListCX.get(CD_index_X) - (Double) ListCX
								.get(CC_index_X))
						+ ","
						+ ((Double) ListMX.get(MD_index_X) - (Double) ListMX
								.get(MC_index_X))
						+ ","
						+ (((Double) ListCX.get(CD_index_X) - (Double) ListCX
								.get(CC_index_X)) - ((Double) ListMX
								.get(MD_index_X) - (Double) ListMX
								.get(MC_index_X))) + "\r\n"
						);
				sCombine = new String();
				listLastAB.clear();
				listLastBC.clear();
				listLastCD.clear();
				listLastDE.clear();
				listLastEF.clear();
				fwprofile.flush();
			}
		}
	}
}
