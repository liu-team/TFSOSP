for (int i = 1; i <= numbInst; i++) {  // ����ÿ��TFRP_JOO����
						for(int r = 1; r <= repli; r++){
							long basis = System.currentTimeMillis();
							double eval_best = Double.MAX_VALUE, eval_sol; //, eval_sol_new;
							int N = Instance[i - 1][0].length;
							int Rep_threshold = 5; // No_impv = 0, 
							int[] Sol = new int[N];
							int[] Sol_new, Sol_new_final;
							double Temp = 0.0;
							
							/*******1. ���������ʼ�⣬��ʼ�¶�******************/
							for(int jj = 0; jj < N; jj++){
								Sol[jj] = random.nextInt(2);
							}
							Temp = 200.0;
							eval_sol = getEvaluationValue(Sol, Instance[i - 1]);
							if(eval_sol < eval_best) {
								eval_best = eval_sol;
							}
							/*******SA�㷨�ĺ���ѭ����������*************************/
							double CPU_Max_Time = N / 10.0;
							long initTime = System.currentTimeMillis();
							long currTime = 0L;
							do {
								for(int rt = 1; rt <= Rep_threshold; rt++){
									/******2. �����ȡ����������������½� ************/
									Sol_new = new int[N];
									Sol_new_final = new int[N];
									// 2.1 ȷ��ȡ������
									int pos1 = 0, pos2 = 0;
									int a = random.nextInt(N);
									int b = random.nextInt(N);
									if(a < b) {
										pos1 = a; 
										pos2 = b;
									} else {
										pos1 = b; 
										pos2 = a;
									}
									// 2.2 ִ������ȡ������
									double eval_reverse = Double.MAX_VALUE;
									double eval_sol_rev = 0.0;
									for(int p = pos1; p <= pos2; p++){
										for(int jj = 0; jj < N; jj++){
											Sol_new[jj] = Sol[jj];
										}
										Sol_new[p] = 1 - Sol[p];
										eval_sol_rev = getEvaluationValue(Sol_new, Instance[i - 1]);
										if(eval_sol_rev < eval_reverse) {
											eval_reverse = eval_sol_rev; 
											for(int jj = 0; jj < N; jj++){
												Sol_new_final[jj] = Sol_new[jj];
											}
										}
									}
									
									// 2.3 ����������ѽ�
									if(eval_reverse < eval_best) {
										eval_best = eval_reverse;
									}
									
									/******3. �����Ƿ����½�ȡ��ԭ�н� ************/
									if(eval_reverse < eval_sol){
										for(int jj = 0; jj < N; jj++){
											Sol[jj] = Sol_new_final[jj];
										}
										eval_sol = eval_reverse;
									}
									if(eval_reverse > eval_sol && Math.random() <= Math.exp(-1 * Arith.div(eval_reverse - eval_sol, Temp))){
										for(int jj = 0; jj < N; jj++){
											Sol[jj] = Sol_new_final[jj];
										}
										eval_sol = eval_reverse;
									}
									/******4. �𽥽����¶�ֵ ************/
									Temp = 0.98 * Temp;
									/******5. ��������δ��øĽ��ĵ������� ************/
									currTime = System.currentTimeMillis();
								}
							} while((currTime - initTime) / 1000.0 <= CPU_Max_Time);
							/************************************************/
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