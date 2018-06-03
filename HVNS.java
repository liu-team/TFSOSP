package tfs_roiop;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;

import basic_arith.Arith;
import jxl.Workbook;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class VNS1Dialog extends JDialog implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 7928675217215859881L;
	JButton jb1, jb2, jb3;
	JComboBox<String> jcb1, jcb2, jcb3, jcb4;
	int[][][] Instance_Set;
	TFSP_Solution JointSol_shaking, JointSol_localsearch;
	int numbInst, N;
	double CRU;
	double[] Opt_obj;
	File file;
	
	public VNS1Dialog(String s1, int[][][] inst,  double rU, double[] opt_obj){
		numbInst = inst.length;
		N = inst[0][0].length;
		Instance_Set = new int[numbInst][4][N];
		Instance_Set = inst;
		CRU = rU;
		Opt_obj = opt_obj;
		
		setLayout(null);
		setTitle(s1);
		setBounds(520, 300, 800, 500);
		
		JLabel j1 = new JLabel("初始解的获取方式");
		j1.setBounds(120, 40, 200, 30);
		add(j1);
		
		String[] str1 = {" JRH ", " JAH "};
		jcb1 = new JComboBox<String>(str1);
		jcb1.setBounds(380, 40, 300, 30);
		add(jcb1);
		
		JLabel j2 = new JLabel("邻域结构的变化次序");
		j2.setBounds(120, 80, 200, 30);
		add(j2);
		
		String[] str2 = {" Exchange(O&I)-->Insert(O2I)-->Insert(I2O) ", 
				 " Exchange(O&I)-->Insert(I2O)-->Insert(O2I) ", 
				 " Insert(O2I)-->Exchange(O&I)-->Insert(I2O) ", 
				 " Insert(O2I)-->Insert(I2O)-->Exchange(O&I) ",
				 " Insert(I2O)-->Exchange(O&I)-->Insert(O2I) ",
				 " Insert(I2O)-->Insert(O2I)-->Exchange(O&I) "};
		jcb2 = new JComboBox<String>(str2);
		jcb2.setBounds(380, 80, 300, 30);
		add(jcb2);
		
		JLabel j3 = new JLabel("参数T的值");
		j3.setBounds(120, 120, 200, 30);
		add(j3);
		
		String[] str3 = {" 0.1 ", " 0.3 ", " 0.5 ", " 0.7 ", " 0.9 "};
		jcb3 = new JComboBox<String>(str3);
		jcb3.setBounds(380, 120, 300, 30);
		add(jcb3);
		
		JLabel j4 = new JLabel("局部搜索算子存在与否");
		j4.setBounds(120, 160, 200, 30);
		add(j4);
		
		String[] str4 = {" With the Best-Improvment operator ", " Without the Best-Improvment operator "};
		jcb4 = new JComboBox<String>(str4);
		jcb4.setBounds(380, 160, 300, 30);
		add(jcb4);
		
		jb1 = new JButton(" 按上述设置运行 ");
		jb1.setBounds(100, 250, 180, 30);
		add(jb1);
		jb2 = new JButton(" 关闭当前对话框 ");
		jb2.setBounds(300, 250, 180, 30);
		add(jb2);
		jb3 = new JButton(" 查看运行结果 ");
		jb3.setBounds(500, 250, 180, 30);
		add(jb3);
		jb3.setEnabled(false);
		
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		jb1.addActionListener(this);
		jb2.addActionListener(this);
		jb3.addActionListener(this);
	}
	public void actionPerformed(ActionEvent e){
		if(e.getSource() == jb2){
			dispose();
		} else if(e.getSource() == jb3){
			try {
				Runtime.getRuntime()
					.exec("C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.exe "
								+ file);
			} catch (IOException ex) {
				System.out.println(ex.toString());
			}
		} else if(e.getSource()==jb1){
			String tr1 = jcb1.getSelectedItem().toString().trim();
			String tr2 = jcb2.getSelectedItem().toString().trim();
			double T = Double.parseDouble(jcb3.getSelectedItem().toString().trim());
			String tr3 = jcb4.getSelectedItem().toString().trim();
			int NS; // Imax;
			int repli = 1;  
			double eval_orig = 0.0, eval_best = 0.0, eval_VNS = 0.0;
			
			WritableWorkbook wwb = null;
			try {
				file = new File(
            			"D:/workspace/MetaHeur_CombOpt/material/Results_VNS_TFSP_JOO.xls");
            	wwb = Workbook.createWorkbook(file);
				if (wwb != null) {
					WritableSheet ws = wwb.createSheet("TFSP_JOO instances by VNS", 0);
            		for (int f = 0; f < 4; f++) {
						ws.setColumnView(f, 25);
					}
            		WritableFont font = new WritableFont(WritableFont.TIMES, 12, WritableFont.BOLD);
					WritableCellFormat format = new WritableCellFormat(font);
					format.setAlignment(jxl.format.Alignment.CENTRE);
					jxl.write.Label cell0, cell1, cell2, cell3;
					cell0 = new jxl.write.Label(0, 0, "Inst Numb", format);
					cell1 = new jxl.write.Label(1, 0, "Rep Numb", format);
					cell2 = new jxl.write.Label(2, 0, "Time Consumed", format);
					cell3 = new jxl.write.Label(3, 0, "Final Obj", format);
					try {
						ws.addCell(cell0);
						ws.addCell(cell1);
						ws.addCell(cell2);
						ws.addCell(cell3);
					} catch (WriteException ex) {
						System.out.println(ex.toString());
					}
					
					for (int i = 1; i <= numbInst; i++) {  // 对于每个TFRP_JOO算例
						for(int r = 1; r <= repli; r++){
							long startTime = System.currentTimeMillis();
							long M = 0L;
							for(int j = 0; j < N; j++){
								M += Instance_Set[i - 1][1][j] + Instance_Set[i - 1][2][j] + Instance_Set[i - 1][3][j];
							}
							/********** 1. 按JAH或者JRH算法，生成一个可行的初始联合决策解  **********/
							TFSP_JOO tsj = new TFSP_JOO();
							if(tr1.equals("JAH")){
								eval_orig = tsj.Job_Addition_Heuristic(Instance_Set[i - 1]);
							} else if(tr1.equals("JRH")){
								eval_orig = tsj.Job_Removal_Heuristic(Instance_Set[i - 1]);
							} 
							
							/********** 2. 执行一系列初始化操作   **********/
							TFSP_Solution JointSol_init = new TFSP_Solution(tsj.bestSol.getInnerSeq(), tsj.bestSol.getOutsourcingSet());
							double Temperature = Arith.mul(T, Arith.div(M, 10 * N));
//							Imax = (int) (5 * Math.sqrt(N) / 1); // N / 4; 
//							int No_impv = 0; 
							eval_VNS = eval_orig;
							eval_best = eval_orig;
							ArrayList<Integer> Seq = new ArrayList<Integer>();
							ArrayList<Integer> Outsourcing = new ArrayList<Integer>();
							
							int[][] JointSolution_best = new int[2][];
							JointSolution_best[0] = new int[JointSol_init.getInnerSeq().size()];
							JointSolution_best[1] = new int[JointSol_init.getOutsourcingSet().size()];
							for(int jj = 0; jj < JointSolution_best[0].length; jj++){
								JointSolution_best[0][jj] = JointSol_init.getInnerSeq().get(jj);
							}
							for(int jj = 0; jj < JointSolution_best[1].length; jj++){
								JointSolution_best[1][jj] = JointSol_init.getOutsourcingSet().get(jj);
							}
							
							int[][] JointSolution_VNS = new int[2][];
							JointSolution_VNS[0] = new int[JointSol_init.getInnerSeq().size()];
							JointSolution_VNS[1] = new int[JointSol_init.getOutsourcingSet().size()];
							for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
								JointSolution_VNS[0][jj] = JointSol_init.getInnerSeq().get(jj);
							}
							for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
								JointSolution_VNS[1][jj] = JointSol_init.getOutsourcingSet().get(jj);
							}
							double eval_shaking = eval_orig, eval_localsearch = eval_orig;
							int[][] InnerSeq_Johnson = new int[3][];
							
							/********** 3. VNS算法的核心循环搜索过程   **********/
							double CPU_Max_Time = N / 10.0;
							long initTime = System.currentTimeMillis();
							long currTime = 0L;
							do {
//							while(No_impv <= Imax){ // 若持续给定次数不改进
//								double eval1 = eval_VNS;
								NS = 1; 
								while(NS <= 3){ // 内部循环
									Seq.clear();
									Outsourcing.clear();
									for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
										Seq.add(JointSolution_VNS[0][jj]);
									}
									for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
										Outsourcing.add(JointSolution_VNS[1][jj]);
									}
									switch(NS){ // 进入抖动(Shaking)阶段
									case 1: 
	                                    if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		exchange_OandI(Seq, Outsourcing, i - 1);
	                                    	}
	                                    if(tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)") || tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		insert_O2I(Seq, Outsourcing, i - 1);
	                                    	}
	                                    if(tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		insert_I2O(Seq, Outsourcing, i - 1);
	                                    	}
	                                    break;
									case 2: 
										if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
											if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
												insert_O2I(Seq, Outsourcing, i - 1);
											}
	                                    if(tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)") || tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		exchange_OandI(Seq, Outsourcing, i - 1);
	                                    	}
	                                    		
	                                    if(tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)") || tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		insert_I2O(Seq, Outsourcing, i - 1);
	                                    	}
	                                    break;
									case 3: 
										if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)"))
											if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
												insert_I2O(Seq, Outsourcing, i - 1);
											}
												
	                                    if(tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)") || tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		insert_O2I(Seq, Outsourcing, i - 1);
	                                    	}
	                                    if(tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
	                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
	                                    		exchange_OandI(Seq, Outsourcing, i - 1);
	                                    	}
	                                    break;
									}
									if(Seq.isEmpty()){
										Outsourcing.clear();
										for(int jj = 0; jj < N; jj++){
											Outsourcing.add(jj + 1);
										}
										double TOC_shaking = 0.0;
										for(int jj = 0; jj < Outsourcing.size(); jj++) { 
											TOC_shaking += Instance_Set[i - 1][3][Outsourcing.get(jj) - 1]; 
										}
										InnerSeq_Johnson = new int[3][0];
										eval_shaking = TOC_shaking;
									} else if(Outsourcing.isEmpty()){
										int Size_seq = N;
								        int[][] CurrSeq = new int[3][Size_seq];
								        InnerSeq_Johnson = new int[3][Size_seq];
								        for(int jj = 0; jj < Size_seq; jj++){
								        	CurrSeq[0][jj] = Instance_Set[i - 1][0][jj];
								        	CurrSeq[1][jj] = Instance_Set[i - 1][1][jj];
								        	CurrSeq[2][jj] = Instance_Set[i - 1][2][jj];
								        }
								        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
								        Seq.clear();
								        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
											Seq.add(InnerSeq_Johnson[0][jj]);
										}
								        eval_shaking = tsj.getMakespan(InnerSeq_Johnson);
									} else {
										double TOC_shaking = 0.0;
										for(int jj = 0; jj < Outsourcing.size(); jj++) { 
											TOC_shaking += Instance_Set[i - 1][3][Outsourcing.get(jj) - 1]; 
										}
										int Size_seq = Seq.size();
								        int[][] CurrSeq = new int[3][Size_seq];
								        InnerSeq_Johnson = new int[3][Size_seq];
								        for(int jj = 0; jj < Size_seq; jj++){
								        	CurrSeq[0][jj] = Seq.get(jj);
								        	CurrSeq[1][jj] = Instance_Set[i - 1][1][CurrSeq[0][jj] - 1];
								        	CurrSeq[2][jj] = Instance_Set[i - 1][2][CurrSeq[0][jj] - 1];
								        }
								        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
								        Seq.clear();
								        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
											Seq.add(InnerSeq_Johnson[0][jj]);
										}
										eval_shaking = tsj.getMakespan(InnerSeq_Johnson) + TOC_shaking;
									}
									JointSol_shaking = new TFSP_Solution(Seq, Outsourcing);
									ArrayList<Integer> Temp_Seq = new ArrayList<Integer>();
									ArrayList<Integer> Temp_Outsourcing = new ArrayList<Integer>();
									for(int j = 0; j < Seq.size(); j++){
										Temp_Seq.add(Seq.get(j));
									}
									for(int j = 0; j < Outsourcing.size(); j++){
										Temp_Outsourcing.add(Outsourcing.get(j));
									}
									
									// Local search 阶段
									if(tr3.equals("With the Best-Improvment operator")){
										switch(NS){ // 进入局部搜索操作
										case 1: 
		                                    if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_exchange_OandI(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    if(tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)") || tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_insert_O2I(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    if(tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_insert_I2O(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    break;
										case 2: 
											if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
												if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
													eval_localsearch = Best_Improvement_insert_O2I(Seq, Outsourcing, eval_shaking, i - 1);
												}
		                                    if(tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)") || tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_exchange_OandI(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    		
		                                    if(tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)") || tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_insert_I2O(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    break;
										case 3: 
											if(tr2.equals("Exchange(O&I)-->Insert(O2I)-->Insert(I2O)") || tr2.equals("Insert(O2I)-->Exchange(O&I)-->Insert(I2O)"))
												if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
													eval_localsearch = Best_Improvement_insert_I2O(Seq, Outsourcing, eval_shaking, i - 1);
												}
													
		                                    if(tr2.equals("Exchange(O&I)-->Insert(I2O)-->Insert(O2I)") || tr2.equals("Insert(I2O)-->Exchange(O&I)-->Insert(O2I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_insert_O2I(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    if(tr2.equals("Insert(O2I)-->Insert(I2O)-->Exchange(O&I)") || tr2.equals("Insert(I2O)-->Insert(O2I)-->Exchange(O&I)"))
		                                    	if(Seq.isEmpty() == false && Outsourcing.isEmpty() == false) {
		                                    		eval_localsearch = Best_Improvement_exchange_OandI(Seq, Outsourcing, eval_shaking, i - 1);
		                                    	}
		                                    break;
										}
										if(eval_localsearch > eval_shaking){
											JointSol_localsearch = new TFSP_Solution(Temp_Seq, Temp_Outsourcing);
											eval_localsearch = eval_shaking;
											// System.out.println(eval_localsearch + " &*^");
										} else {
											JointSol_localsearch = new TFSP_Solution(Seq, Outsourcing);
											// System.out.println(eval_localsearch + " &*^");
										}
										// 更新eval_best的值
										if(eval_localsearch < eval_best) {
											eval_best = eval_localsearch;
											JointSolution_best[0] = new int[JointSol_localsearch.getInnerSeq().size()];
											JointSolution_best[1] = new int[JointSol_localsearch.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_best[0].length; jj++){
												JointSolution_best[0][jj] = JointSol_localsearch.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_best[1].length; jj++){
												JointSolution_best[1][jj] = JointSol_localsearch.getOutsourcingSet().get(jj);
											}
											
											System.out.println("inhouse:" + Arrays.toString(JointSolution_best[0]));
											System.out.println("Outsource:" + Arrays.toString(JointSolution_best[1]));
											System.out.println(eval_localsearch + "***");
										}
										// 进入 Move or not 阶段
										if(eval_localsearch < eval_VNS){  
											eval_VNS = eval_localsearch;
											NS = 1;
											JointSolution_VNS[0] = new int[JointSol_localsearch.getInnerSeq().size()];
											JointSolution_VNS[1] = new int[JointSol_localsearch.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
												JointSolution_VNS[0][jj] = JointSol_localsearch.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
												JointSolution_VNS[1][jj] = JointSol_localsearch.getOutsourcingSet().get(jj);
											}
										} else if(eval_localsearch > eval_VNS && Math.random() <= Math.exp(Arith.mul(-1, Arith.div(eval_localsearch - eval_VNS, Temperature)))){
											eval_VNS = eval_localsearch;
											NS = 1;
											System.out.println(">>>>>>");
											JointSolution_VNS[0] = new int[JointSol_localsearch.getInnerSeq().size()];
											JointSolution_VNS[1] = new int[JointSol_localsearch.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
												JointSolution_VNS[0][jj] = JointSol_localsearch.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
												JointSolution_VNS[1][jj] = JointSol_localsearch.getOutsourcingSet().get(jj);
											}
										} else {
											NS = NS + 1;
										}
									} else {
										// 更新 eval_best的值
										if(eval_shaking < eval_best) {
											eval_best = eval_shaking;
											JointSolution_best[0] = new int[JointSol_shaking.getInnerSeq().size()];
											JointSolution_best[1] = new int[JointSol_shaking.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_best[0].length; jj++){
												JointSolution_best[0][jj] = JointSol_shaking.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_best[1].length; jj++){
												JointSolution_best[1][jj] = JointSol_shaking.getOutsourcingSet().get(jj);
											}
											
											System.out.println("inhouse:" + Arrays.toString(JointSolution_best[0]));
											System.out.println("Outsource:" + Arrays.toString(JointSolution_best[1]));
											System.out.println(eval_shaking + "***");
										}
										// 进入 Move or not 阶段
										if(eval_shaking < eval_VNS){  
											eval_VNS = eval_shaking;
											NS = 1;
											JointSolution_VNS[0] = new int[JointSol_shaking.getInnerSeq().size()];
											JointSolution_VNS[1] = new int[JointSol_shaking.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
												JointSolution_VNS[0][jj] = JointSol_shaking.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
												JointSolution_VNS[1][jj] = JointSol_shaking.getOutsourcingSet().get(jj);
											}
											
										} else if(eval_shaking > eval_VNS && Math.random() <= Math.exp(Arith.mul(-1, Arith.div(eval_shaking - eval_VNS, Temperature)))){
											eval_VNS = eval_shaking;
											NS = 1;
											System.out.println(">>>>>>");
											JointSolution_VNS[0] = new int[JointSol_shaking.getInnerSeq().size()];
											JointSolution_VNS[1] = new int[JointSol_shaking.getOutsourcingSet().size()];
											for(int jj = 0; jj < JointSolution_VNS[0].length; jj++){
												JointSolution_VNS[0][jj] = JointSol_shaking.getInnerSeq().get(jj);
											}
											for(int jj = 0; jj < JointSolution_VNS[1].length; jj++){
												JointSolution_VNS[1][jj] = JointSol_shaking.getOutsourcingSet().get(jj);
											}
										} else {
											NS = NS + 1;
										}
									}
								}
								// 内部子循环的结束
								// 核算未获得改进的次数
//								double eval2 = eval_VNS;
//								if(eval2 >= eval1) 
//									No_impv = No_impv + 1;
//								else 
//									No_impv = 0;
								currTime = System.currentTimeMillis();
							} while(eval_best > Opt_obj[i - 1]);
							//while((currTime - initTime) / 1000.0 <= CPU_Max_Time);// 外部循环结束
							
							long endTime=System.currentTimeMillis();
							jxl.write.Number number1 = new jxl.write.Number(0,
									repli * (i - 1) + r, i, format);
							jxl.write.Number number2 = new jxl.write.Number(1,
									repli * (i - 1) + r, r, format);
							jxl.write.Number number3 = new jxl.write.Number(2,
									repli * (i - 1) + r, (endTime - startTime) / 1000.0, format);
							jxl.write.Number number4 = new jxl.write.Number(3,
									repli * (i - 1) + r, eval_best, format);
//							cell0 = new jxl.write.Label(0, i, String.valueOf(i), format);
//							cell1 = new jxl.write.Label(1, i, String.valueOf(Arith.sub(endTime, startTime) / 1000.0), format);
//							cell2 = new jxl.write.Label(2, i, String.valueOf(eval_best), format);
							
							try {
								ws.addCell(number1);
								ws.addCell(number2);
								ws.addCell(number3);
								ws.addCell(number4);
							} catch (WriteException ex) {
								System.out.println(ex.toString());
							}
						}
					}
					wwb.write();
                    wwb.close();
                    file.deleteOnExit();
				}
			} catch (IOException ex) {
            	System.out.println(ex.toString());
            } catch (WriteException ex) {
            	System.out.println(ex.toString());
            }
			jb3.setEnabled(true);
		}
	}
	
	public void exchange_OandI(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, int index){ // 邻域结构 exchange_O&I
		Collections.shuffle(Seq);
        Collections.shuffle(Outsourcing);
        int a = (Integer) Seq.get(0); 
        int b = (Integer) Outsourcing.get(0); 
        int removeIndex = -1;
		for(int ii = 0; ii < Seq.size(); ii++){
			if(a == Seq.get(ii)) {
				removeIndex = ii;
			}
		}
		Seq.remove(removeIndex);
		removeIndex = -1;
		for(int ii = 0; ii < Outsourcing.size(); ii++){
			if(b == Outsourcing.get(ii)) {
				removeIndex = ii;
			}
		}
        Outsourcing.remove(removeIndex);
        Seq.add(b);
        Outsourcing.add(a);
	}
	
	public void insert_O2I(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, int index) {
		Collections.shuffle(Outsourcing);
		int b = (Integer) Outsourcing.get(0);
		Seq.add(b);
		Outsourcing.remove(0);
	}
	
	public void insert_I2O(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, int index) {
		Collections.shuffle(Seq);
		int a = (Integer) Seq.get(0); 
		Outsourcing.add(a);
		Seq.remove(0);
	}
	
	public double Best_Improvement_exchange_OandI(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, double eval_shaking, int index){  
		int Seq_size = Seq.size();
		int Outsourcing_size = Outsourcing.size();
		TFSP_JOO tsj = new TFSP_JOO();
		double eval_localsearch = 0.0, eval_best = eval_shaking;
		int[][] InnerSeq_Johnson;
		ArrayList<Integer> Init_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Init_Outsourcing = new ArrayList<Integer>();
		ArrayList<Integer> Best_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Best_Outsourcing = new ArrayList<Integer>();
		for(int j = 0; j < Seq.size(); j++){
			Best_Seq.add(Seq.get(j));
		}
		for(int j = 0; j < Outsourcing.size(); j++){
			Best_Outsourcing.add(Outsourcing.get(j));
		}
		
		// 找出具备最大外包成本的外包工件的所在位置
		int pos = 0, toc = Integer.MIN_VALUE, job;
		for(int j = 0; j < Outsourcing.size(); j++){
			job = Outsourcing.get(j);
			if(Instance_Set[index][3][job - 1] > toc) {
				toc = Instance_Set[index][3][job - 1];
				pos = j;
			}
		}
		
		for(int i = 0; i < Seq_size; i++){
			Init_Seq.clear();
			Init_Outsourcing.clear();
			for(int j = 0; j < Seq.size(); j++){
				Init_Seq.add(Seq.get(j));
			}
			for(int j = 0; j < Outsourcing.size(); j++){
				Init_Outsourcing.add(Outsourcing.get(j));
			}
			
			int a = (Integer) Init_Seq.get(i);
			int b = (Integer) Init_Outsourcing.get(pos); 
			Init_Seq.remove(i);
			Init_Outsourcing.remove(pos);
			Init_Seq.add(b);
			Init_Outsourcing.add(a);
			
			double TOC_localsearch = 0.0;
			for(int jj = 0; jj < Init_Outsourcing.size(); jj++) { 
				TOC_localsearch += Instance_Set[index][3][Init_Outsourcing.get(jj) - 1]; 
			}
			int Init_Seq_Size = Init_Seq.size();
	        int[][] CurrSeq = new int[3][Init_Seq_Size];
	        InnerSeq_Johnson = new int[3][Init_Seq_Size];
	        for(int jj = 0; jj < Init_Seq_Size; jj++){
	        	CurrSeq[0][jj] = Init_Seq.get(jj);
	        	CurrSeq[1][jj] = Instance_Set[index][1][CurrSeq[0][jj] - 1];
	        	CurrSeq[2][jj] = Instance_Set[index][2][CurrSeq[0][jj] - 1];
	        }
	        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
	        Init_Seq.clear();
	        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
				Init_Seq.add(InnerSeq_Johnson[0][jj]);
			}
			eval_localsearch = tsj.getMakespan(InnerSeq_Johnson) + TOC_localsearch;
			
			// System.out.println(eval_localsearch + "  局部搜索的解");
			if(eval_localsearch < eval_best){
				eval_best = eval_localsearch;
				Best_Seq.clear();
				Best_Outsourcing.clear();
				for(int j = 0; j < Init_Seq.size(); j++){
					Best_Seq.add(Init_Seq.get(j));
				}
				for(int j = 0; j < Init_Outsourcing.size(); j++){
					Best_Outsourcing.add(Init_Outsourcing.get(j));
				}
			}
		}
		
		// 找出具备最大加工时间之和的内部工件的所在位置
		pos = 0;
		toc = Integer.MIN_VALUE;
		for(int j = 0; j < Seq.size(); j++){
			job = Seq.get(j);
			if(Instance_Set[index][1][job - 1] + Instance_Set[index][2][job - 1] > toc) {
				toc = Instance_Set[index][1][job - 1] + Instance_Set[index][2][job - 1];
				pos = j;
			}
		}
		
		for(int i = 0; i < Outsourcing_size; i++){
			Init_Seq.clear();
			Init_Outsourcing.clear();
			for(int j = 0; j < Seq.size(); j++){
				Init_Seq.add(Seq.get(j));
			}
			for(int j = 0; j < Outsourcing.size(); j++){
				Init_Outsourcing.add(Outsourcing.get(j));
			}
			
			int a = (Integer) Init_Seq.get(pos);
			int b = (Integer) Init_Outsourcing.get(i); 
			Init_Seq.remove(pos);
			Init_Outsourcing.remove(i);
			Init_Seq.add(b);
			Init_Outsourcing.add(a);
			
			double TOC_localsearch = 0.0;
			for(int jj = 0; jj < Init_Outsourcing.size(); jj++) { 
				TOC_localsearch += Instance_Set[index][3][Init_Outsourcing.get(jj) - 1]; 
			}
			int Init_Seq_Size = Init_Seq.size();
	        int[][] CurrSeq = new int[3][Init_Seq_Size];
	        InnerSeq_Johnson = new int[3][Init_Seq_Size];
	        for(int jj = 0; jj < Init_Seq_Size; jj++){
	        	CurrSeq[0][jj] = Init_Seq.get(jj);
	        	CurrSeq[1][jj] = Instance_Set[index][1][CurrSeq[0][jj] - 1];
	        	CurrSeq[2][jj] = Instance_Set[index][2][CurrSeq[0][jj] - 1];
	        }
	        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
	        Init_Seq.clear();
	        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
				Init_Seq.add(InnerSeq_Johnson[0][jj]);
			}
			eval_localsearch = tsj.getMakespan(InnerSeq_Johnson) + TOC_localsearch;
			
			// System.out.println(eval_localsearch + "  局部搜索的解");
			if(eval_localsearch < eval_best){
				eval_best = eval_localsearch;
				Best_Seq.clear();
				Best_Outsourcing.clear();
				for(int j = 0; j < Init_Seq.size(); j++){
					Best_Seq.add(Init_Seq.get(j));
				}
				for(int j = 0; j < Init_Outsourcing.size(); j++){
					Best_Outsourcing.add(Init_Outsourcing.get(j));
				}
			}
		}
		
		Seq.clear();
		Outsourcing.clear();
		for(int j = 0; j < Best_Seq.size(); j++){
			Seq.add(Best_Seq.get(j));
		}
		for(int j = 0; j < Best_Outsourcing.size(); j++){
			Outsourcing.add(Best_Outsourcing.get(j));
		}
		return eval_best;
	}
	
	public double Best_Improvement_insert_O2I(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, double eval_shaking, int index) {
		int Outsource_size = Outsourcing.size();
		TFSP_JOO tsj = new TFSP_JOO();
		double eval_localsearch = 0.0, eval_best = eval_shaking;
		int[][] InnerSeq_Johnson;
		ArrayList<Integer> Init_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Init_Outsourcing = new ArrayList<Integer>();
		ArrayList<Integer> Best_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Best_Outsourcing = new ArrayList<Integer>();
		for(int j = 0; j < Seq.size(); j++){
			Best_Seq.add(Seq.get(j));
		}
		for(int j = 0; j < Outsourcing.size(); j++){
			Best_Outsourcing.add(Outsourcing.get(j));
		}
		
		for(int i = 0; i < Outsource_size; i++){
			Init_Seq.clear();
			Init_Outsourcing.clear();
			for(int j = 0; j < Seq.size(); j++){
				Init_Seq.add(Seq.get(j));
			}
			for(int j = 0; j < Outsourcing.size(); j++){
				Init_Outsourcing.add(Outsourcing.get(j));
			}
			int b = (Integer) Init_Outsourcing.get(i);
			Init_Seq.add(b);
			Init_Outsourcing.remove(i);
			
			if(Init_Outsourcing.isEmpty()){
				int Init_Seq_Size = N;
		        int[][] CurrSeq = new int[3][Init_Seq_Size];
		        InnerSeq_Johnson = new int[3][Init_Seq_Size];
		        for(int jj = 0; jj < Init_Seq_Size; jj++){
		        	CurrSeq[0][jj] = Instance_Set[index][0][jj];
		        	CurrSeq[1][jj] = Instance_Set[index][1][jj];
		        	CurrSeq[2][jj] = Instance_Set[index][2][jj];
		        }
		        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
		        Init_Seq.clear();
		        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
					Init_Seq.add(InnerSeq_Johnson[0][jj]);
				}
		        eval_localsearch = tsj.getMakespan(InnerSeq_Johnson);
			} else {
				double TOC_localsearch = 0.0;
				for(int jj = 0; jj < Init_Outsourcing.size(); jj++) { 
					TOC_localsearch += Instance_Set[index][3][Init_Outsourcing.get(jj) - 1]; 
				}
				int Init_Seq_Size = Init_Seq.size();
		        int[][] CurrSeq = new int[3][Init_Seq_Size];
		        InnerSeq_Johnson = new int[3][Init_Seq_Size];
		        for(int jj = 0; jj < Init_Seq_Size; jj++){
		        	CurrSeq[0][jj] = Init_Seq.get(jj);
		        	CurrSeq[1][jj] = Instance_Set[index][1][CurrSeq[0][jj] - 1];
		        	CurrSeq[2][jj] = Instance_Set[index][2][CurrSeq[0][jj] - 1];
		        }
		        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
		        Init_Seq.clear();
		        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
					Init_Seq.add(InnerSeq_Johnson[0][jj]);
				}
				eval_localsearch = tsj.getMakespan(InnerSeq_Johnson) + TOC_localsearch;
			}
			// System.out.println(eval_localsearch + "  局部搜索的解");
			if(eval_localsearch < eval_best){
				eval_best = eval_localsearch;
				Best_Seq.clear();
				Best_Outsourcing.clear();
				for(int j = 0; j < Init_Seq.size(); j++){
					Best_Seq.add(Init_Seq.get(j));
				}
				for(int j = 0; j < Init_Outsourcing.size(); j++){
					Best_Outsourcing.add(Init_Outsourcing.get(j));
				}
			}
		}
		Seq.clear();
		Outsourcing.clear();
		for(int j = 0; j < Best_Seq.size(); j++){
			Seq.add(Best_Seq.get(j));
		}
		for(int j = 0; j < Best_Outsourcing.size(); j++){
			Outsourcing.add(Best_Outsourcing.get(j));
		}
		return eval_best;
	}
	
	public double Best_Improvement_insert_I2O(ArrayList<Integer> Seq, ArrayList<Integer> Outsourcing, double eval_shaking, int index) {
		int Seq_size = Seq.size();
		TFSP_JOO tsj = new TFSP_JOO();
		double eval_localsearch = 0.0, eval_best = eval_shaking;
		int[][] InnerSeq_Johnson;
		ArrayList<Integer> Init_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Init_Outsourcing = new ArrayList<Integer>();
		ArrayList<Integer> Best_Seq = new ArrayList<Integer>();
		ArrayList<Integer> Best_Outsourcing = new ArrayList<Integer>();
		for(int j = 0; j < Seq.size(); j++){
			Best_Seq.add(Seq.get(j));
		}
		for(int j = 0; j < Outsourcing.size(); j++){
			Best_Outsourcing.add(Outsourcing.get(j));
		}
		
		for(int i = 0; i < Seq_size; i++){
			Init_Seq.clear();
			Init_Outsourcing.clear();
			for(int j = 0; j < Seq.size(); j++){
				Init_Seq.add(Seq.get(j));
			}
			for(int j = 0; j < Outsourcing.size(); j++){
				Init_Outsourcing.add(Outsourcing.get(j));
			}
			int a = (Integer) Init_Seq.get(i);
			Init_Outsourcing.add(a);
			Init_Seq.remove(i);
			
			if(Init_Seq.isEmpty()){
				Init_Outsourcing.clear();
				for(int jj = 0; jj < N; jj++){
					Init_Outsourcing.add(jj + 1);
				}
				double TOC_localsearch = 0.0;
				for(int jj = 0; jj < Init_Outsourcing.size(); jj++) { 
					TOC_localsearch += Instance_Set[index][3][Init_Outsourcing.get(jj) - 1]; 
				}
				InnerSeq_Johnson = new int[3][0];
				eval_localsearch = TOC_localsearch;
			} else {
				double TOC_localsearch = 0.0;
				for(int jj = 0; jj < Init_Outsourcing.size(); jj++) { 
					TOC_localsearch += Instance_Set[index][3][Init_Outsourcing.get(jj) - 1]; 
				}
				int Init_Seq_Size = Init_Seq.size();
		        int[][] CurrSeq = new int[3][Init_Seq_Size];
		        InnerSeq_Johnson = new int[3][Init_Seq_Size];
		        for(int jj = 0; jj < Init_Seq_Size; jj++){
		        	CurrSeq[0][jj] = Init_Seq.get(jj);
		        	CurrSeq[1][jj] = Instance_Set[index][1][CurrSeq[0][jj] - 1];
		        	CurrSeq[2][jj] = Instance_Set[index][2][CurrSeq[0][jj] - 1];
		        }
		        InnerSeq_Johnson = tsj.getOptOriginalSchedule(CurrSeq);
		        Init_Seq.clear();
		        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
					Init_Seq.add(InnerSeq_Johnson[0][jj]);
				}
				eval_localsearch = tsj.getMakespan(InnerSeq_Johnson) + TOC_localsearch;
			}
			// System.out.println(eval_localsearch + "  局部搜索的解");
			if(eval_localsearch < eval_best){
				eval_best = eval_localsearch;
				Best_Seq.clear();
				Best_Outsourcing.clear();
				for(int j = 0; j < Init_Seq.size(); j++){
					Best_Seq.add(Init_Seq.get(j));
				}
				for(int j = 0; j < Init_Outsourcing.size(); j++){
					Best_Outsourcing.add(Init_Outsourcing.get(j));
				}
			}
		}
		Seq.clear();
		Outsourcing.clear();
		for(int j = 0; j < Best_Seq.size(); j++){
			Seq.add(Best_Seq.get(j));
		}
		for(int j = 0; j < Best_Outsourcing.size(); j++){
			Outsourcing.add(Best_Outsourcing.get(j));
		}
		return eval_best;
	}
	
}