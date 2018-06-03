WritableWorkbook wwb = null;
			int[][] InnerSeq_Johnson = new int[3][];
			
			try {
				File file = new File(
            			"D:/workspace/MetaHeur_CombOpt/material/Results_ImplicitEnum_TFSP_JOO.xls");
            	wwb = Workbook.createWorkbook(file);
            	
            	if (wwb != null) {
            		WritableSheet ws = wwb.createSheet("TFSP_JOO by Implicit Enum", 0);
            		for (int f = 0; f < 3; f++) {
						ws.setColumnView(f, 30);
					}
            		WritableFont font = new WritableFont(WritableFont.TIMES, 12, WritableFont.BOLD);
					WritableCellFormat format = new WritableCellFormat(font);
					format.setAlignment(jxl.format.Alignment.CENTRE);
					jxl.write.Label cell0, cell1, cell2;
					cell0 = new jxl.write.Label(0, 0, "Number of Inhouse Jobs", format);
					cell1 = new jxl.write.Label(1, 0, "Time Consumed", format);
					cell2 = new jxl.write.Label(2, 0, "Optimum", format);
					try {
						ws.addCell(cell0);
						ws.addCell(cell1);
						ws.addCell(cell2);
					} catch (WriteException ex) {
						System.out.println(ex.toString());
					}
            		
					for (int i = 1; i <= numbInst; i++) {  // 对于每个TFRP_JOO算例
						long basis = System.currentTimeMillis();
						long current = 0L;
						double eval_best = Double.MAX_VALUE, eval = 0.0;
						int number_innerjob_best = 0;
						int[] A = new int[N];
						
						ArrayList<Integer> Seq = new ArrayList<Integer>();
						ArrayList<Integer> Outsourcing = new ArrayList<Integer>();
						
						while(isAllOne(A) == false && (current - basis) / 1000.0 <= 50000.0){
							/********** 1. 将当前二进制解转换为二元组解(1代表被外包出去；0代表在内部完成加工)**********/
							Seq.clear();
							Outsourcing.clear();
							for(int jj = 0; jj < A.length; jj++){
								if(A[jj] == 1) Outsourcing.add(Instance[i - 1][0][jj]);
								else Seq.add(Instance[i - 1][0][jj]);
							}
							/********** 2. 计算出各个二元组解的评价值**********/
							if(Outsourcing.isEmpty()){
								int Size_seq = N;
						        int[][] CurrSeq = new int[3][Size_seq];
						        InnerSeq_Johnson = new int[3][Size_seq];
						        for(int jj = 0; jj < Size_seq; jj++){
						        	CurrSeq[0][jj] = Instance[i - 1][0][jj];
						        	CurrSeq[1][jj] = Instance[i - 1][1][jj];
						        	CurrSeq[2][jj] = Instance[i - 1][2][jj];
						        }
						        InnerSeq_Johnson = getOptOriginalSchedule(CurrSeq);
						        Seq.clear();
						        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
									Seq.add(InnerSeq_Johnson[0][jj]);
								}
						        eval = getMakespan(InnerSeq_Johnson);
							} else {
								double TOC_shaking = 0.0;
								for(int jj = 0; jj < Outsourcing.size(); jj++) { 
									TOC_shaking += Instance[i - 1][3][Outsourcing.get(jj) - 1]; 
								}
								int Size_seq = Seq.size();
						        int[][] CurrSeq = new int[3][Size_seq];
						        InnerSeq_Johnson = new int[3][Size_seq];
						        for(int jj = 0; jj < Size_seq; jj++){
						        	CurrSeq[0][jj] = Seq.get(jj);
						        	CurrSeq[1][jj] = Instance[i - 1][1][CurrSeq[0][jj] - 1];
						        	CurrSeq[2][jj] = Instance[i - 1][2][CurrSeq[0][jj] - 1];
						        }
						        InnerSeq_Johnson = getOptOriginalSchedule(CurrSeq);
						        Seq.clear();
						        for(int jj = 0; jj < InnerSeq_Johnson[0].length; jj++){
									Seq.add(InnerSeq_Johnson[0][jj]);
								}
								eval = getMakespan(InnerSeq_Johnson) + TOC_shaking;
							}
							/******* 3. 不断更新最佳评价值*******/
							if(eval < eval_best) {
								eval_best = eval;
								number_innerjob_best = InnerSeq_Johnson[0].length;
							}
							/******************************/
			        		int j = N - 1;
			                while(A[j] == 1 && j >= 0) {
			                	A[j] = 0;
			                	j = j - 1;
			                }
			                A[j] = A[j] + 1;
			                current = System.currentTimeMillis();
						}
						if(isAllOne(A) == true) {
							Outsourcing.clear();
							for(int jj = 0; jj < N; jj++){
								Outsourcing.add(jj + 1);
							}
							double TOC_shaking = 0.0;
							for(int jj = 0; jj < Outsourcing.size(); jj++) { 
								TOC_shaking += Instance[i - 1][3][Outsourcing.get(jj) - 1]; 
							}
							InnerSeq_Johnson = new int[3][0];
							eval = TOC_shaking;
							if(eval < eval_best) {
								eval_best = eval;
								number_innerjob_best = InnerSeq_Johnson[0].length;
							}
							current = System.currentTimeMillis();
							cell0 = new jxl.write.Label(0, i, String.valueOf(number_innerjob_best), format);
							cell1 = new jxl.write.Label(1, i, String.valueOf(Arith.sub(current, basis) / 1000.0), format);
							cell2 = new jxl.write.Label(2, i, String.valueOf(eval_best), format);
							try {
								ws.addCell(cell0);
								ws.addCell(cell1);
								ws.addCell(cell2);
							} catch (WriteException ex) {
								System.out.println(ex.toString());
							}
						}
					} // 结束对各个算例的循环求解 
            		wwb.write();
                    wwb.close();
                    file.deleteOnExit();
            	}
			} catch (IOException ex) {
            	System.out.println(ex.toString());
            } catch (WriteException ex) {
            	System.out.println(ex.toString());
            }