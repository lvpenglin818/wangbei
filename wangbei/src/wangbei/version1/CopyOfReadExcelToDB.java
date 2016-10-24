package wangbei.version1;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * ���������ڽ����ϴ���excel�ļ���excel�ļ��е�������ʤͨ��HMI����ϸ��Ϣ��excel�ļ��ĸ�ʽ����ϸ�ĵ�˵�� Ȼ�����е�����д�����ݿ���
 * 
 * @author LPL
 * 
 */
public class CopyOfReadExcelToDB {
	public static void main(String[] args) throws Exception, Exception {
		// File fileC = new File(
		// "C://Users//lvpenglin//Desktop//c.xls");
		// File fileM = new File(
		// "C://Users//lvpenglin//Desktop//m.xls");
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

	public static Map readExcel(File file) throws BiffException, IOException {
		// File file = new File("filePath");
		Workbook wb = Workbook.getWorkbook(file);// ���ļ�����ȡ��Excel����������
		// ��ʼ����Excel������
		Map map = new LinkedHashMap<String, List>();
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			// ���λ�ȡ�������е�excel���
			Sheet sheet = wb.getSheet(sheetIndex);
			// System.out.println("��һ��sheet���У�" + sheet.getRows() + "��"
			// + sheet.getColumns() + "��");
			double[] f = null;
			StringBuilder sb = new StringBuilder();
			// ����excel����е�ÿһ�У��Զ���ӵڶ��п�ʼ������е�һ��Ϊ��ͷ
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
							// System.out.println(sb);
							map.put(sb.toString(), list);
							sb = new StringBuilder();
							f = null;
							j = 0;
							break;
						} else {
							temp = sheet.getCell(0, i).getContents();
							double f_temp = Double.parseDouble(temp
									.substring(temp.indexOf(" ") + 1));// ��ȡ�ӿո�ʼ�ĺ������е��ַ�
							list.add(f_temp);
						}
					}
				}
				if (i == sheet.getRows() - 1) {// �ж��Ƿ�λ���������һ��
					break;
				}
			}
		}
		System.out.println(map.size());
		return map;
	}

	public static Map readExcelToXValue(File file) throws BiffException,
			IOException {
		// File file = new File("filePath");
		Workbook wb = Workbook.getWorkbook(file);// ���ļ�����ȡ��Excel����������
		// ��ʼ����Excel������
		Map mapX = new LinkedHashMap<String, List>();
		for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
			// ���λ�ȡ�������е�excel���
			Sheet sheet = wb.getSheet(sheetIndex);
			// System.out.println("��һ��sheet���У�" + sheet.getRows() + "��"
			// + sheet.getColumns() + "��");
			double[] f = null;
			StringBuilder sb = new StringBuilder();
			// ����excel����е�ÿһ�У��Զ���ӵڶ��п�ʼ������е�һ��Ϊ��ͷ
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
									0, temp.indexOf(";")));// ��ȡ�ӵ�һ���ַ���ʼ���ֺ�֮ǰ�����е��ַ�
							listX.add(f_tempX);
						}
					}
				}
				if (i == sheet.getRows() - 1) {// �ж��Ƿ�λ���������һ��
					break;
				}
			}
		}
		System.out.println(mapX.size());
		return mapX;
	}

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

		// ����Ҫд���excel���
		WritableWorkbook wbOut = Workbook.createWorkbook(new File(
				"C://Users//lvpenglin//Desktop//������.xls"));
		// ͨ��Excel�ļ���ȡ��һ��������sheet
		WritableSheet sheet1 = wbOut.createSheet("sheet1", 0);
		WritableSheet sheet2 = wbOut.createSheet("sheet2", 2);
		int pddrows = 3;
		int profilerows = 3;
		int pddCol = 0;
		int profileCol = 0;

		FileWriter fwprofile = new FileWriter(new File(
				"C://Users//lvpenglin//Desktop//���profile.csv"));
		FileWriter fwpdd = new FileWriter(new File(
				"C://Users//lvpenglin//Desktop//���pdd.csv"));
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
				// sheet1.addCell(new Label(pddCol++,pddrows,sCombine));
				// sheet1.addCell(new
				// Label(pddCol++,pddrows,Deviation.getMean((Double[])
				// listLastLeft.toArray(new Double[listLastLeft.size()]))));
				// sheet1.addCell(new
				// Label(pddCol++,pddrows,Deviation.ComputeVariance2((Double[])
				// listLastLeft.toArray(new Double[listLastLeft.size()]))));
				// sheet1.addCell(new
				// Label(pddCol++,pddrows,Deviation.getMean((Double[])
				// listLastRight.toArray(new Double[listLastLeft.size()]))));
				// sheet1.addCell(new
				// Label(pddCol++,pddrows++,Deviation.ComputeVariance2((Double[])
				// listLastRight.toArray(new Double[listLastLeft.size()]))));
				fwpdd.write(sCombine
						+ ","
						+ Deviation.ComputeVariance((Double[]) listLastLeft
								.toArray(new Double[listLastLeft.size()]))
						+","
						+ Deviation.ComputeVariance((Double[]) listLastRight
								.toArray(new Double[listLastRight.size()]))
						+ "\r\n");
				// fw.write(sCombine
				// + "���ľ�ֵ�ͱ�׼�����������"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastLeft
				// .toArray(new Double[listLastLeft.size()]))
				// + "\r\n"
				// + sCombine
				// + "�Ҳ�ľ�ֵ�ͱ�׼�����������"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastRight
				// .toArray(new Double[listLastRight.size()]))
				// + "\r\n");
				fwpdd.flush();
				System.out.println(sCombine + "���ľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastLeft
								.toArray(new Double[listLastLeft.size()])));
				System.out.println(sCombine + "�Ҳ�ľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastRight
								.toArray(new Double[listLastRight.size()])));
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

//				double A = (midValue * 0.1);
				double A = (double) ListC.get(0);

				double B = (midValue * 0.2);

				double C = (midValue * 0.8);

				double D = (midValue * 0.8);

				double E = (midValue * 0.2);

//				double F = (midValue * 0.1);
				double F = (double) ListC.get(ListC.size()-1);

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
							MA_index_X = i + 1;// A����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
						}
						if ((Double) ListM.get(i) < MB_indexValue
								&& (Double) ListM.get(i + 1) >= MB_indexValue) {
							MB_index_X = i + 1;
						}
						if ((Double) ListC.get(i) < CA_indexValue
								&& (Double) ListC.get(i + 1) >= CA_indexValue) {
							CA_index_X = i + 1;// A����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
						}
						if ((Double) ListC.get(i) < CB_indexValue
								&& (Double) ListC.get(i + 1) >= CB_indexValue) {
							CB_index_X = i + 1;
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
							MC_index_X = i + 1;// C����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
						}
						if ((Double) ListM.get(i) > MD_indexValue
								&& (Double) ListM.get(i + 1) <= MD_indexValue) {
							MD_index_X = i + 1;// D����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
						}
						if ((Double) ListC.get(i) > CC_indexValue
								&& (Double) ListC.get(i + 1) <= CC_indexValue) {
							CC_index_X = i + 1;// C����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
						}
						if ((Double) ListC.get(i) > CD_indexValue
								&& (Double) ListC.get(i + 1) <= CD_indexValue) {
							CD_index_X = i + 1;// D����M�е�index����Ҫ�����Ӧ��x���꣬x�����ֵӦ�õ����洢��mapX��
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
												/ (Double) ListM.get(i));
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
						+ ((Double) ListCX.get(CC_index_X) - (Double) ListCX
								.get(CB_index_X))
								+ ","
						+ ((Double) ListMX.get(MC_index_X) - (Double) ListMX
								.get(MB_index_X))
						+ ","
						+ (((Double) ListCX.get(CC_index_X) - (Double) ListCX
								.get(CB_index_X)) - ((Double) ListMX
								.get(MC_index_X) - (Double) ListMX
								.get(MB_index_X)))
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
				// fw.write(sCombine
				// + "\r\n"
				// + "MA_index_XΪ��"
				// + MA_index_X
				// + "MB_index_XΪ��"
				// + MB_index_X
				// + "MC_index_XΪ��"
				// + MC_index_X
				// + "MD_index_XΪ��"
				// + MD_index_X
				// + "\r\n"
				// + "CA_index_XΪ��"
				// + CA_index_X
				// + "CB_index_XΪ��"
				// + CB_index_X
				// + "CC_index_XΪ��"
				// + CC_index_X
				// + "CD_index_XΪ��"
				// + CD_index_X
				// + "\r\n"
				// + "Measured��AB֮��X�ľ���Ϊ��"
				// + ((Double) ListMX.get(MB_index_X) - (Double) ListMX
				// .get(MA_index_X))
				// + "\r\n"
				// + "Measured��BC֮��X�ľ���Ϊ��"
				// + ((Double) ListMX.get(MC_index_X) - (Double) ListMX
				// .get(MB_index_X))
				// + "\r\n"
				// + "Measured��CD֮��X�ľ���Ϊ��"
				// + ((Double) ListMX.get(MD_index_X) - (Double) ListMX
				// .get(MC_index_X))
				// + "\r\n"
				// + "Computed��AB֮��X�ľ���Ϊ��"
				// + ((Double) ListCX.get(CB_index_X) - (Double) ListCX
				// .get(CA_index_X))
				// + "\r\n"
				// + "Computed��BC֮��X�ľ���Ϊ��"
				// + ((Double) ListCX.get(CC_index_X) - (Double) ListCX
				// .get(CB_index_X))
				// + "\r\n"
				// + "Computed��CD֮��X�ľ���Ϊ��"
				// + ((Double) ListCX.get(CD_index_X) - (Double) ListCX
				// .get(CC_index_X))
				// + "\r\n"
				// + "AB��ֵΪ��"
				// + (((Double) ListCX.get(CB_index_X) - (Double) ListCX
				// .get(CA_index_X)) - ((Double) ListMX
				// .get(MB_index_X) - (Double) ListMX
				// .get(MA_index_X)))
				// + "\r\n"
				// + "BC��ֵΪ��"
				// + (((Double) ListCX.get(CC_index_X) - (Double) ListCX
				// .get(CB_index_X)) - ((Double) ListMX
				// .get(MC_index_X) - (Double) ListMX
				// .get(MB_index_X)))
				// + "\r\n"
				// + "CD��ֵΪ��"
				// + (((Double) ListCX.get(CD_index_X) - (Double) ListCX
				// .get(CC_index_X)) - ((Double) ListMX
				// .get(MD_index_X) - (Double) ListMX
				// .get(MC_index_X)))
				// + "\r\n"
				// + "AB�εľ�ֵ�ͱ�׼�����������:"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastAB
				// .toArray(new Double[listLastAB.size()]))
				// + "\r\n"
				// + sCombine
				// + "\r\n"
				// + "BC�εľ�ֵ�ͱ�׼�����������:"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastBC
				// .toArray(new Double[listLastBC.size()]))
				// + "\r\n"
				// + sCombine
				// + "\r\n"
				// + "CD�εľ�ֵ�ͱ�׼�����������:"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastCD
				// .toArray(new Double[listLastCD.size()]))
				// + "\r\n"
				// + sCombine
				// + "\r\n"
				// + "DE�εľ�ֵ�ͱ�׼�����������:"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastDE
				// .toArray(new Double[listLastDE.size()]))
				// + "\r\n"
				// + sCombine
				// + "\r\n"
				// + "EF�εľ�ֵ�ͱ�׼�����������:"
				// + "\r\n"
				// + Deviation.ComputeVariance((Double[]) listLastEF
				// .toArray(new Double[listLastEF.size()]))
				// + "\r\n");
				System.out.println();
				System.out.println(sCombine + "AB�εľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastAB
								.toArray(new Double[listLastAB.size()])));
				System.out.println();
				System.out.println(sCombine + "BC�εľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastBC
								.toArray(new Double[listLastBC.size()])));
				System.out.println();
				System.out.println(sCombine + "CD�εľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastCD
								.toArray(new Double[listLastCD.size()])));
				System.out.println();
				System.out.println(sCombine + "DE�εľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastDE
								.toArray(new Double[listLastDE.size()])));
				System.out.println();
				System.out.println(sCombine + "EF�εľ�ֵ�ͱ�׼�����������");
				System.out.println(Deviation
						.ComputeVariance((Double[]) listLastEF
								.toArray(new Double[listLastEF.size()])));
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
