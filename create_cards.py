import openpyxl as xl
import re



def create_cards(filename):
    wb = xl.load_workbook(filename, data_only=True)
    sheets_names = wb.sheetnames

    sheets_names = [x for x in sheets_names if re.match(r"LC[0-9]+", x) ]

    for worksheets in wb.sheetnames:
        worksheet = wb[worksheets]
        if worksheet.title in sheets_names:
            text_file = open(f"Output/{worksheet.title}.txt", "a+")
            LC_id = re.compile(r"[^LC0].*")
            LC_id = LC_id.findall(worksheet.title)[0]
            LC_id = 100 * int(LC_id)
             
            #GRID INERTIAL 
            node_grid_inertial = 3800001 +LC_id
            for row in range(36, 42): #worksheet.max_row
                x_inertial = worksheet.cell(row,2).value
                y_inertial = worksheet.cell(row,3).value
                z_inertial = worksheet.cell(row,4).value

           
                grid_string_inertial= f"GRID,{node_grid_inertial},0,{x_inertial},{y_inertial},{z_inertial}"
        
                if (row == 36):
                    text_file.write(f"BULKHEAD__GRID__INTERTIAL for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(grid_string_inertial)
                text_file.write("\n")
                node_grid_inertial +=1

            text_file.write("\n")
            
            #FORCE_INERTIAL
            node_f_inertial = 3800001 + LC_id
            for row in range(36, 42): #worksheet.max_row
                Fx_inertial = worksheet.cell(row,5).value
                Fy_inertial = worksheet.cell(row,6).value
                Fz_inertial = worksheet.cell(row,7).value
                force_id = re.compile(r"[^LC0].*")
                force_id = force_id.findall(worksheet.title)[0]
                force_string_inertial = f"FORCE,{force_id},{node_f_inertial},0,1.0,{Fx_inertial},{Fy_inertial},{Fz_inertial}"

                if (row == 36):
                    text_file.write(f"BULKHEAD__FORCE__INTERTIAL for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(force_string_inertial)
                text_file.write("\n")
                node_f_inertial += 1

            text_file.write("\n")

            #GRID AERO 
            node_grid_aero = node_grid_inertial  
            for row in range(36, 42): #worksheet.max_row
                x_aero = worksheet.cell(row,2).value
                y_aero = worksheet.cell(row,8).value
                z_aero = worksheet.cell(row,9).value

                grid_string_aero= f"GRID,{node_grid_aero},0,{x_aero},{y_aero},{z_aero}"


                if (row == 36):
                    text_file.write(f"BULKHEAD__GRID__AERO for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(grid_string_aero)
                text_file.write("\n")
                node_grid_aero += 1    

            text_file.write("\n")

            # FORCE_AERO
            node_f_aero = node_f_inertial 
            for row in range(36, 42): #worksheet.max_row
                Fx_aero = worksheet.cell(row,10).value
                Fy_aero = worksheet.cell(row,11).value
                Fz_aero = worksheet.cell(row,12).value

                force_id = re.compile(r"[^LC0].*")
                force_id = force_id.findall(worksheet.title)[0]
                force_string_aero = f"FORCE,{force_id},{node_f_aero},0,1.0,{Fx_aero},{Fy_aero},{Fz_aero}"

                if (row == 36):
                    text_file.write(f"BULKHEAD__FORCE__AERO for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(force_string_aero)
                text_file.write("\n")
                node_f_aero += 1    

            text_file.write("\n")            

            # HT_VT_GRID_INERTIAL
            ht_vt_grid_node_intertial = 3900001 + LC_id
            for row in range(7,10):
                # MOMENT,MomentID,Node,0,1.0,Mx,My,Mz
                x_inertial =  worksheet.cell(row, 15).value   
                y_inertial =  worksheet.cell(row, 16).value   
                z_inertial =  worksheet.cell(row, 17).value   
                
                grid_string_inertial= f"GRID,{ht_vt_grid_node_intertial},0,{x_inertial},{y_inertial},{z_inertial}"
                if (row == 7):
                    text_file.write(f"HT-VT__GRID_INERTIAL for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(grid_string_inertial)
                text_file.write("\n")
                
                ht_vt_grid_node_intertial +=1      

            text_file.write("\n")

            # HT_VT_FORCE_INERTIAL
            ht_vt_force_node_intertial = 3900001 + LC_id
            for row in range(7,10):
                # MOMENT,MomentID,Node,0,1.0,Mx,My,Mz
                Fx_inertial = worksheet.cell(row,18).value
                Fy_inertial = worksheet.cell(row,19).value
                Fz_inertial = worksheet.cell(row,20).value
                force_id = re.compile(r"[^LC0].*")
                force_id = force_id.findall(worksheet.title)[0]
                force_string_inertial = f"FORCE,{force_id},{ht_vt_force_node_intertial},0,1.0,{Fx_inertial},{Fy_inertial},{Fz_inertial}"
    
                if (row == 7):
                    text_file.write(f"HT-VT__FORCE__INERTIAL for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(force_string_inertial)
                text_file.write("\n")
                
                ht_vt_force_node_intertial +=1
            
            text_file.write("\n")

            # HT_VT_MOMENT_INERTIAL
            ht_vt_mom_node_inertial = 3900001 + LC_id
            for row in range(7,10):
                # MOMENT,MomentID,Node,0,1.0,Mx,My,Mz
                Mx = worksheet.cell(row,21).value
                My = worksheet.cell(row,22).value
                Mz = worksheet.cell(row,23).value
                moment_id = re.compile(r"[^LC0].*")
                moment_id = moment_id.findall(worksheet.title)[0]

                moment_string_inertial = f"MOMENT,{moment_id},{ht_vt_mom_node_inertial},0,1.0,{Mx},{My},{Mz}"
    
                if (row == 7):
                    text_file.write(f"HT-VT__MOMENT__INERTIAL for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(moment_string_inertial)
                text_file.write("\n")
                
                ht_vt_mom_node_inertial +=1    

            text_file.write("\n")

            # HT_VT_GRID_AERO
            ht_vt_grid_node_aero = ht_vt_grid_node_intertial 
            for row in range(7,10):
                # MOMENT,MomentID,Node,0,1.0,Mx,My,Mz
                x_aero = worksheet.cell(row,24).value
                y_aero = worksheet.cell(row,25).value
                z_aero = worksheet.cell(row,26).value

                grid_string_aero= f"GRID,{ht_vt_grid_node_aero},0,{x_aero},{y_aero},{z_aero}"

              
                if (row == 7):
                    text_file.write(f"HT-VT__GRID_AERO for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(grid_string_aero)
                text_file.write("\n")
                
                ht_vt_grid_node_aero += 1             
            
            text_file.write("\n")
            
            # HT_VT_FORCE_INERTIAL
            ht_vt_force_node_aero = ht_vt_force_node_intertial 
            for row in range(7,10):
                # MOMENT,MomentID,Node,0,1.0,Mx,My,Mz
                Fx_aero = worksheet.cell(row,27).value
                Fy_aero = worksheet.cell(row,28).value
                Fz_aero = worksheet.cell(row,29).value
                force_id = re.compile(r"[^LC0].*")
                force_id = force_id.findall(worksheet.title)[0]
                force_string_aero = f"FORCE,{force_id},{ht_vt_force_node_aero},0,1.0,{Fx_aero},{Fy_aero},{Fz_aero}"
    
                if (row == 7):
                    text_file.write(f"HT-VT__FORCE__AERO for {worksheet.title}")
                    text_file.write("\n")
                text_file.write(force_string_aero)
                text_file.write("\n")
                
                ht_vt_force_node_aero +=1

            text_file.write("\n")

            # VF 
            VF_node = 3950001 + LC_id
            text_file.write(f"VF__GRID__INERTIAL for {worksheet.title}")
            text_file.write("\n")            
            x_inertial =  worksheet.cell(19, 15).value   
            y_inertial =  worksheet.cell(19, 16).value   
            z_inertial =  worksheet.cell(19, 17).value   
            VF_grid_inertial = f"GRID,{VF_node},0,{x_inertial},{y_inertial},{z_inertial}"
            text_file.write(VF_grid_inertial)
            text_file.write("\n")

            text_file.write("\n")

            text_file.write(f"VF__FROCE_INERTIAL for {worksheet.title}")
            text_file.write("\n")            
            Fx_inertial = worksheet.cell(19,18).value
            Fy_inertial = worksheet.cell(19,19).value
            Fz_inertial = worksheet.cell(19,20).value
            force_id = re.compile(r"[^LC0].*")
            force_id = force_id.findall(worksheet.title)[0]
            VF_force_inertial = f"FORCE,{force_id},{VF_node},0,1.0,{Fx_inertial},{Fy_inertial},{Fz_inertial}"
            text_file.write(VF_force_inertial)
            text_file.write("\n")

            text_file.write("\n")

            text_file.write(f"VF__MOMENT__INERTIAL for {worksheet.title}")
            text_file.write("\n")            
            Mx = worksheet.cell(19,21).value
            My = worksheet.cell(19,22).value
            Mz = worksheet.cell(19,23).value
            moment_id = re.compile(r"[^LC0].*")
            moment_id = moment_id.findall(worksheet.title)[0]
            moment_VF_inertial = f"MOMENT,{moment_id},{VF_node},0,1.0,{Mx},{My},{Mz}"
            text_file.write(moment_VF_inertial)
            text_file.write("\n")

            text_file.write("\n")

            text_file.write(f"VF__GRID AERO for {worksheet.title}")
            text_file.write("\n")            
            x_aero =  worksheet.cell(19, 24).value   
            y_aero =  worksheet.cell(19, 25).value   
            z_aero =  worksheet.cell(19, 26).value   
            VF_grid_aero = f"GRID,{VF_node+1},0,{x_aero},{y_aero},{z_aero}"
            text_file.write(VF_grid_aero)
            text_file.write("\n")

            text_file.write("\n")

            text_file.write(f"VF__FORCE_AERO for {worksheet.title}")
            text_file.write("\n")            
            Fx_aero = worksheet.cell(19,27).value
            Fy_aero = worksheet.cell(19,28).value
            Fz_aero = worksheet.cell(19,29).value      
            VF_force_aero = f"FORCE,{force_id},{VF_node+1},0,1.0,{Fx_aero},{Fy_aero},{Fz_aero}"
            text_file.write(VF_force_aero)
            text_file.write("\n")
             

            text_file.close()



create_cards(r"input.xlsx")