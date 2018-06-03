for (int i = 1; i <= numbInst; i++) {  // ����ÿ��TFRP_JOO����
						for(int r = 1; r <= repli; r++){
							long basis = System.currentTimeMillis();
							double eval_best = Double.MAX_VALUE; //, eval_sol_new;
							int N = Instance[i - 1][0].length;
							int HMS = N;
							double HMCR = 0.95, PAR = 0.05;
							int[][] HM = new int[HMS][N];
							int[] HV = new int[N];
							double[] eval_HM = new double[HMS];
							double eval_HV = 0.0;
							
							/*******1. ��ʼ�����������  ******************/
							for(int h = 0; h < HMS; h++){
								for(int jj = 0; jj < N; jj++){
									HM[h][jj] = random.nextInt(2);
								}
							}
							double CPU_Max_Time = N / 10.0;
							long initTime = System.currentTimeMillis();
							long currTime = 0L;
							int h_best = 0;
							
							/*******HS�㷨�ĺ���ѭ����������*************************/
							do {
								/*******2. ���۵�ǰ�������и���������Ӧֵ ******************/
								for(int h = 0; h < HMS; h++){
									eval_HM[h] = getEvaluationValue(HM[h], Instance[i - 1]);
									/*******��ʱ�����������ֵ*******/
									if(eval_HM[h] < eval_best) {
										eval_best = eval_HM[h];
										h_best = h;
									}
								}
								/******* 3. ���ݵ�ǰ��HM��������һ���º�������*******/
								for(int g = 0; g < N; g++){
									if(Math.random() < HMCR){ // ����˼������
//										int rand = random.nextInt(HMS);
										HV[g] = HM[h_best][g];
										if(Math.random() < PAR){ // ����΢������
											HV[g] = 1 - HV[g]; 
										}
									} else { // ���ѡȡ����
										HV[g] = random.nextInt(2);
									}
								}
								/******* 4. ���º��������*******/
								eval_HV = getEvaluationValue(HV, Instance[i - 1]);
								double eval_max = Double.MIN_VALUE;
								int H = 0;
								for(int h = 0; h < HMS; h++){
									if(eval_HM[h] > eval_max){
										eval_max = eval_HM[h];
										H = h;
									}
								}
//								if(eval_HV <= eval_max){
								eval_HM[H] = eval_HV;
								for(int g = 0; g < N; g++){
									HM[H][g] = HV[g];
								}
//								}
								currTime = System.currentTimeMillis();
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