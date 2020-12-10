# ExcelSheetManipulation
These programs create Grid Cards, Force cards and Moment cards in the from of .txt files from excel sheets.<br/>
The .txt files can be used as an input to Nastran for load calculation.<br/>
Grid Cards have the format  --> GRID,Node,0,x,y,z<br/>
Force Cards have the format --> FORCE,ForceID,Node,0,1.0,Fx,Fy,Fz<br/>
Moment Cards have the format --> MOMENT,MomentID,Node,0,1.0,Mx,My,Mz<br/>
