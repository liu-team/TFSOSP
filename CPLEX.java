WritableWorkbook wwb = null;
			IloCplex cplex;
			long basis = 0L, current;
			double running;
			
			int[] pt1, pt2, oc;
			try {
				File file = new File(
            			"D:/workspace/MetaHeur_CombOpt/material/Results_CPLEX_TFSP_JOO.xls");
            	wwb = Workbook.createWorkbook(file);
            	if (wwb != null) {
            		WritableSheet ws = wwb.createSheet("TFSP_JOO instances by CPLEX", 0);
            		for (int f = 0; f < 3; f++) {
						ws.setColumnView(f, 20);
					}
            		WritableFont font = new WritableFont(WritableFont.TIMES, 12, WritableFont.BOLD);
					WritableCellFormat format = new WritableCellFormat(font);
					format.setAlignment(jxl.format.Alignment.CENTRE);
					jxl.write.Label cell0, cell1, cell2;
					cell0 = new jxl.write.Label(0, 0, "Solution Status", format);
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
						/********* 对该算例运用CPLEX进行求解**********/
						pt1 = new int[N];
						pt2 = new int[N];
	                    oc = new int[N];
	                    for(int w = 0; w < N; w++){
                        	pt1[w] = Instance[i - 1][1][w];
                        	pt2[w] = Instance[i - 1][2][w];
                        	oc[w] = Instance[i - 1][3][w];
                        }
	                    cplex = new IloCplex();
	                    /********** 1. 声明决策变量**********/
	                    IloNumVar[] y = new IloNumVar[N];
                        y = cplex.boolVarArray(N);
                        
                        IloNumVar[][] x = new IloNumVar[N][];
                        for(int jj = 0; jj < N; jj++)
                            x[jj] = cplex.boolVarArray(N);
                        
                        IloNumVar[] W = new IloNumVar[N];
                        double[] lb = new double[N];
                        for(int jj = 0; jj < N; jj++) {
                        	lb[jj] = 0.0;
                        }
                		double[] ub = new double[N]; 
                		for(int jj = 0; jj < N; jj++) {
	                        ub[jj] = Double.MAX_VALUE;
	                    } // 是否需要在此对变量W进行进一步的限制???
                		for(int jj = 0; jj < N; jj++) {
                			W[jj] = cplex.numVar(lb[jj], ub[jj], IloNumVarType.Int);
                		}
                		
                		/********** 2. 构造目标函数**********/
                		IloNumExpr obj1 = cplex.numExpr();
                		IloNumExpr[] obj0 = new IloNumExpr[N];
	                    for(int kk = 0; kk < N; kk++) {
	                        obj0[kk] = cplex.linearNumExpr();
	                        obj0[kk] = cplex.sum(obj0[kk], W[kk]);
	                        for(int jj = 0; jj < N; jj++){
	                        	obj0[kk] = cplex.sum(obj0[kk], cplex.prod(pt2[jj], x[jj][kk]));
	                        }
	                        for(int hh = 0; hh <= kk; hh++){
	                        	for(int jj = 0; jj < N; jj++){
	                        		obj0[kk] = cplex.sum(obj0[kk], cplex.prod(pt1[jj], x[jj][hh]));
	                        	}
	                        }
	                    }
	                    obj1 = cplex.max(obj0);
	                    IloLinearNumExpr obj2 = cplex.linearNumExpr();
	                    obj2 = cplex.scalProd(oc, y);
	                    IloNumExpr objective = cplex.linearNumExpr();
	                    objective = cplex.sum(obj1, obj2);
                        cplex.addObjective(IloObjectiveSense.Minimize, objective);
                        /********** 3. 添加约束条件**********/
                        for(int kk = 0; kk < N; kk++){ // 3.1  设置约束条件1
                        	IloNumExpr con1 = cplex.linearNumExpr();
                        	for(int jj = 0; jj < N; jj++){
                        		con1 = cplex.sum(con1, x[jj][kk]);
                        	}
                        	cplex.addLe(con1, 1); 
                        }
                        
                        IloNumExpr con11 = cplex.linearNumExpr();
                        for(int jj = 0; jj < N; jj++){
                        	for(int kk = 0; kk < N; kk++){
                        		con11 = cplex.sum(con11, x[jj][kk]);
                        	}
                        }
                        IloNumExpr con12 = cplex.linearNumExpr();
                        for(int jj = 0; jj < N; jj++){
                        	con12 = cplex.sum(con12, cplex.sum(1, cplex.prod(y[jj], -1)));
                        }
                        cplex.addEq(con11, con12); // 3.2 设置约束条件2
                        
                        for(int jj = 0; jj < N; jj++){ // 3.3 设置约束条件3
                        	IloNumExpr con2 = cplex.linearNumExpr();
                        	for(int kk = 0; kk < N; kk++){
                        		con2 = cplex.sum(con2, x[jj][kk]);
                        	}
                        	cplex.addEq(con2, cplex.sum(1, cplex.prod(y[jj], -1)));
                        }
                        
                        IloNumExpr con31, con32;
                        for(int kk = 1; kk < N; kk++){ // 3.3 设置约束条件3
                        	con31 = cplex.linearNumExpr();
                        	con32 = cplex.linearNumExpr();
                        	
                        	con31 = cplex.sum(con31, W[kk]);
                        	for(int jj = 0; jj < N; jj++) {
                        		con31 = cplex.sum(con31, cplex.prod(pt1[jj], x[jj][kk])); 
                        	}
                        	
                        	con32 = cplex.sum(con32, W[kk - 1]);
                        	for(int jj = 0; jj < N; jj++) {
                        		con32 = cplex.sum(con32, cplex.prod(pt2[jj], x[jj][kk - 1]));
                        	}
                        	cplex.addGe(con31, con32);
                        }
                        
                        cplex.addEq(W[0], 0); // 3.5 设置约束条件5
                        
                        for(int kk = 1; kk < N; kk++){ // 3.6 设置约束条件6
                        	cplex.addGe(W[kk], 0);
                        }
                        
                        cplex.setParam(IloCplex.DoubleParam.TiLim, 10000);
                        cplex.setParam(IloCplex.IntParam.Threads, 1);
                        basis = System.currentTimeMillis();
                        if (cplex.solve()){
                        	current = System.currentTimeMillis();
                        	running = Arith.sub(current, basis);
                        	cell0 = new jxl.write.Label(0, i, String.valueOf(cplex.getStatus()), format);
                        	cell1 = new jxl.write.Label(1, i, String.format("%.3f", running / 1000.0), format);
                        	cell2 = new jxl.write.Label(2, i, String.valueOf(cplex.getObjValue()), format);
                        	try{
                        		ws.addCell(cell0);
                        		ws.addCell(cell1);
                        		ws.addCell(cell2);
                        	}catch(WriteException ee){
                        		ee.printStackTrace();
                        	}
                        }
                        cplex.end();
					}
            		wwb.write();
                    wwb.close();
                    file.deleteOnExit();
            	}
			} catch(IloException ie){
            	System.err.println("Concert exception cautht: "+ ie);
            } catch (IOException ex) {
            	System.out.println(ex.toString());
            } catch (WriteException ex) {
            	System.out.println(ex.toString());
            }