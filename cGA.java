WritableWorkbook wwb = null;
			Random random = new Random();
			int repli = 30;
			
			try {
				File file = new File(
            			"D:/workspace/MetaHeur_CombOpt/material/Results_GA_TFSP_JOO.xls");
            	wwb = Workbook.createWorkbook(file);
            	
            	if (wwb != null) {
            		WritableSheet ws = wwb.createSheet("TFSP_JOO instances by GA", 0);
            		for (int f = 0; f < 4; f++) {
						ws.setColumnView(f, 30);
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
							long basis = System.currentTimeMillis();
							double eval_best = Double.MAX_VALUE;
							
							int N = Instance[i - 1][0].length;
							int pop_size = N;
							int[][] Pop = new int[pop_size][N];
							int[][] new_Pop = new int[3 * pop_size][N];
							int new_point = 0;
							double[] eval_Pop = new double[pop_size];
							double[] eval_new_Pop = new double[3 * pop_size];
							double prob_c = 0.90, prob_m = 0.10;
							
							/*******1. 初始化种群 ******************/
							for(int p = 0; p < pop_size; p++){
								for(int jj = 0; jj < N; jj++){
									Pop[p][jj] = random.nextInt(2);
								}
							}
							double CPU_Max_Time = N / 10.0;
							long initTime = System.currentTimeMillis();
							long currTime = 0L;
							/*******GA算法的核心循环搜索过程*************************/
							do {
								/*******2. 评价当前种群的适应值 ******************/
								for(int p = 0; p < pop_size; p++){
									eval_Pop[p] = getEvaluationValue(Pop[p], Instance[i - 1]);
									/*******及时更新最佳评价值*******/
									if(eval_Pop[p] < eval_best) {
										eval_best = eval_Pop[p];
									}
								}
								/******* 3. 单点交叉操作,将产生的新解置于新的种群中*******/
								new_point = 0; 
								for(int cro = 0; cro < pop_size / 2; cro++){
									if(Math.random() <= prob_c){
										int[] crossover1 = new int[N];
										int[] crossover2 = new int[N];
										// 3.1 随机选取出2个参与交叉操作的染色体
										int rand1 = random.nextInt(pop_size);
										int rand2 = random.nextInt(pop_size);
										for(int jj = 0; jj < N; jj++){
											crossover1[jj] = Pop[rand1][jj];
											crossover2[jj] = Pop[rand2][jj];
										}
										// 3.2 确定出单点交叉的位置
										int cro_pos = random.nextInt(N);
										// 3.3 执行单点交叉操作
										int temp = 0;
										for(int jj = cro_pos; jj < N; jj++){
											temp = crossover1[jj];
											crossover1[jj] = crossover2[jj];
											crossover2[jj] = temp;
										}
										// 3.4 将两个新染色体置于新的种群中
										for(int jj = 0; jj < N; jj++){
											new_Pop[new_point][jj] = crossover1[jj];
										}
										new_point++;
										for(int jj = 0; jj < N; jj++){
											new_Pop[new_point][jj] = crossover2[jj];
										}
										new_point++;
									}
								}
								/******* 4. 双点取反变异操作,将产生的新解置于新的种群中*******/
								for(int mut = 0; mut < pop_size; mut++){
									if(Math.random() <= prob_m){
										int[] mutation = new int[N];
										// 4.1 确定变异的双点
										int mut_pos1 = 0, mut_pos2 = 0;
										int a = random.nextInt(N);
										int b = random.nextInt(N);
										if(a < b) {
											mut_pos1 = a; 
											mut_pos2 = b;
										} else {
											mut_pos1 = b; 
											mut_pos2 = a;
										}
										// 4.2  执行双点取反变异操作
										for(int jj = 0; jj < mut_pos1; jj++){
											mutation[jj] = Pop[mut][jj];
										}
										mutation[mut_pos1] = 1 - Pop[mut][mut_pos1];
										for(int jj = mut_pos1 + 1; jj < mut_pos2; jj++){
											mutation[jj] = Pop[mut][jj];
										}
										mutation[mut_pos2] = 1 - Pop[mut][mut_pos2];
										for(int jj = mut_pos2 + 1; jj < N; jj++){
											mutation[jj] = Pop[mut][jj];
										}
										// 4.3 将新染色体置于新种群中
										for(int jj = 0; jj < N; jj++){
											new_Pop[new_point][jj] = mutation[jj];
										}
										new_point++;
									}
								}
								
								/******5. 3-联赛选择操作, 从新种群中选择个体组成下一代种群************/
								// 5.1 完整组建新种群，补充上当代种群
								for(int p = 0; p < pop_size; p++){
									for(int jj = 0; jj < N; jj++){
										new_Pop[new_point + p][jj] = Pop[p][jj];
									}
								}
								new_point = new_point + pop_size;
								// 5.2 评价新的种群中各个染色体的适应值
								for(int new_p = 0; new_p < new_point; new_p++){
									eval_new_Pop[new_p] = getEvaluationValue(new_Pop[new_p], Instance[i - 1]);
									/*******及时更新最佳评价值*******/
									if(eval_new_Pop[new_p] < eval_best) {
										eval_best = eval_new_Pop[new_p];
									}
								}
								// 5.3 选择生成下一代的种群
								for(int p = 0; p < pop_size; p++){
									int a = random.nextInt(new_point);
									int b = random.nextInt(new_point);
									int c = random.nextInt(new_point); 
									double curr_eval = 0.0;
									if(eval_new_Pop[a] <= eval_new_Pop[b]) {
										for(int jj = 0; jj < N; jj++){
											Pop[p][jj] = new_Pop[a][jj];
										}
										curr_eval = eval_new_Pop[a];
									} else {
										for(int jj = 0; jj < N; jj++){
											Pop[p][jj] = new_Pop[b][jj];
										}
										curr_eval = eval_new_Pop[b];
									}
									if(eval_new_Pop[c] <= curr_eval){
										for(int jj = 0; jj < N; jj++){
											Pop[p][jj] = new_Pop[c][jj];
										}
									}
								}
								
								currTime = System.currentTimeMillis();
							} while((currTime - initTime) / 1000.0 <= CPU_Max_Time);
							long current = System.currentTimeMillis();
							jxl.write.Number number1 = new jxl.write.Number(0,
									repli * (i - 1) + r, i , format);
							jxl.write.Number number2 = new jxl.write.Number(1,
									repli * (i - 1) + r, r, format);
							jxl.write.Number number3 = new jxl.write.Number(2,
									repli * (i - 1) + r, (current - basis) / 1000.0 , format);
							jxl.write.Number number4 = new jxl.write.Number(3,
									repli * (i - 1) + r, eval_best, format);
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