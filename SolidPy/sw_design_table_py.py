# from attr import dataclass
from dataclasses import dataclass, field
from operator import index
import swtoolkit as swtk
import win32com.client
import os
import csv
import pythoncom
import pywintypes





def check_file_exists(file_path):
    """Checks if a file exists. Returns True if it does, False if it doesn't."""
    if os.path.isfile(file_path):
        return 
    else:
        open(file_path, "w").close

def get_design_table_from_model_as_list(sw_model):
    """Returns the design table from a model. Returns None if no design table exists."""
    design_table = sw_model.GetDesignTable

    if design_table is None:
        return None
    else:
        dt = []
        nTotRow = design_table.GetTotalRowCount
        nTotCol = design_table.GetTotalColumnCount
        nTotRow = nTotRow + 1
        nTotCol = nTotCol + 1


        design_table.Attach

        # Generate Header Row
        header_row = []
        header_row.append("Config_Name")

        for c_index in range(1, nTotCol):
            header_row.append(design_table.GetHeaderText(c_index-1))

        dt.append(header_row)

        # Populate List of table row data
        for r_index in range(1, nTotRow):
            row_data = []
            for c_index in range(nTotCol):
                row_data.append(design_table.GetEntryValue(r_index, c_index))
            dt.append(row_data)    
        
        design_table.Detach

        return dt


@dataclass
class SW_Configuration:
    """Class for Solidworks Configurations - represents a row in a design table"""
    config_name: str
    header_columns: list
    row_data: list

    def update_config(self):
        """Updates the configuration values in Solidworks Design Table"""
        pass

    def add_config(self):
        """Adds a new configuration to the Solidworks Design Table"""
        pass


@dataclass
class SW_DesignTable:
    """Class for Solidworks Design Tables"""
    sw_model: win32com.client.CDispatch
    sw_template_dir: str 
    name: str = field(init=False)
    table: list = field(init=False)

    def __post_init__(self):
        self.name = self.sw_model.GetTitle
        self.table = self.get_design_table_from_model_as_list()

    def get_design_table_from_model_as_list(self):
        """Returns the design table from a model. Returns None if no design table exists."""
        design_table = self.sw_model.GetDesignTable
        if design_table is None:
            return None
        else:
            bool = design_table.Attach

            # Generate Header Row
            header_row = []
            header_row.append("Config_Name")
            nTotRow = design_table.GetTotalRowCount + 1
            nTotCol = design_table.GetTotalColumnCount + 1
            for c_index in range(1, nTotCol):
                header_row.append(design_table.GetHeaderText(c_index-1))
            print(f"Header Row: {header_row}")
                
            print(f"Table Data:")
            for r_index in range( nTotRow):
                row_data = []
                for c_index in range(nTotCol):
                    print(f"      Row: {r_index} | Col: {c_index} | Value: {design_table.GetEntryText(r_index, c_index)}")
                    row_data.append(design_table.GetEntryText(r_index, c_index))
                print(f"      Row Data: {row_data}")
                # configs.append(SW_Configuration(row_data[0], header_row, row_data[1:]))  

            # Populate List of SW_Configuration objects
            configs = []
            for r_index in range(1, nTotRow):
                row_data = []
                for c_index in range(nTotCol):
                    row_data.append(design_table.GetEntryText(r_index, c_index))
                print(f"      Row Data: {row_data}")
                configs.append(SW_Configuration(row_data[0], header_row, row_data[1:]))    

            bool = design_table.Detach
            return configs

    def update_table(self):
        """Updates the configuration values in Solidworks Design Table"""
        design_table = self.sw_model.GetDesignTable
        design_table.Attach
        nTableRows = design_table.GetTotalRowCount
        print(f"nTableRows: {nTableRows}")

        for r_idx, config_row in enumerate(self.table):
            for dt_row in range(nTableRows):
                # Get the configuration name from the design table
                sw_config_name = design_table.GetEntryValue(dt_row, 0)
                print(f"sw_config_name: {sw_config_name}")
                # If the configuration names match, update the row
                if sw_config_name == config_row.config_name:
                    for c_idx, value in enumerate(config_row.row_data):
                        is_text = isinstance(value, str)
                        design_table.SetEntryValue(r_idx, c_idx+1, is_text, value)  # c_idx+1 because we're starting from second column

        design_table.UpdateTable(2, True)  # 2 corresponds to swUpdateDesignTableAll constant
        design_table.Detach

    def add_config(self, config_row: SW_Configuration):
        """Adds a new configuration to the Solidworks Design Table"""
        design_table = self.sw_model.GetDesignTable
        num_cols = design_table.GetTotalColumnCount
        num_cols = num_cols + 1
        # empty_col_array = 
        # cell_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [])
        # design_table.Attach
        cells = [config_row.config_name] + config_row.row_data
        # cells = config_row.row_data
        print(f"\n\n{cells}\n\n")
        cell_array = []
        for index, cell in enumerate(cells):
            print(f"      Cell | Type: {cell} | {type(cell)}")
            cell_array.append(cell)
        cell_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, cell_array)
        
        # for cell in cells:
        #     print(f"      Cell | Type: {cell} | {type(cell)}")

        # for cell in cells:
        #     cell = str(cell)

        # for cell in cells:
        #     print(f"      Cell | Type: {cell} | {type(cell)}")

        # Here is an example of creating a VARIANT array
        print(f"      New Row Cells: {cells} | {type(cells)}")
        print(f"      New Row Cells: {cell_array} | {type(cell_array)}")
        worksheet = design_table.EditTable
        print(f"      Worksheet: {worksheet}")
        # boolstatus = win32com.client.VARIANT(pythoncom.VT_BOOL, design_table.AddRow(cell_array))
        print(cell_array.value)
        try:
            boolstatus = design_table.AddRow(cell_array)
            print(f"AddRow successful?: {boolstatus}")
            boolstatus = design_table.UpdateTable(2, True)  # 2 corresponds to swUpdateDesignTableAll constant
            print(f"UpdateTable successful?: {boolstatus}")
        except Exception as e:
            print(f"Error in AddRow or UpdateTable: {e}")
        # design_table.Detach


    def write_to_csv(self):
        """Writes the design table to a csv file with the same name as the model. Creates new CSV if doesnt exist, completely overwrites if it does."""
        csv_file_path = os.path.join(self.sw_template_dir, self.name + ".csv")
        try:
            with open(csv_file_path, 'w') as csv_file:
                for config in self.table:
                    csv_file.write(",".join([config.config_name] + config.row_data) + "\n")
        except FileNotFoundError:
            open(csv_file_path, "w").close()

    def update_from_csv(self):
        """Updates the design table from a CSV file"""
        print(f"\n\nUpdating design table from CSV file...\n")
        # Build the path to the CSV file
        csv_file_path = os.path.join(self.sw_template_dir, self.name + ".csv")

        # Read the CSV file
        with open(csv_file_path, 'r') as csv_file:
            reader = csv.reader(csv_file)
            header_row = next(reader)
            print(f"  Header Row: {header_row}")
            csv_configurations = []

            for row in reader:
                # Create a SW_Configuration for each row in the CSV file
                csv_configurations.append(SW_Configuration(row[0], header_row[1:], row[1:]))
                print(f"  CSV Config: {csv_configurations[-1]}")
        for config in csv_configurations:
            print(f"CSV Config: {config}")
        # Iterate over the configurations from the CSV file
        print(f"\n\nTable Object: {self.table}\n\n")
        for csv_config in csv_configurations:
            # Check if this configuration is already in the design table
            in_table = False
            for dt_config in self.table:
                print(f"  DT Config: {dt_config}")
                print(f"  Comparing {dt_config.config_name} to {csv_config.config_name}")
                if dt_config.config_name == csv_config.config_name:
                    in_table = True
                    # The configuration is in the design table, so update it
                    dt_config.row_data = csv_config.row_data
                    # self.update_config(dt_config)
                    # break
                # else:
                #     dt_config.row_data = csv_config.row_data
            if not in_table:
                # The configuration is not in the design table, so add it
                print(f"  Adding {csv_config} to design table...")
                self.add_config(csv_config)
        self.update_table()
            


sw = swtk.SolidWorks()

# get active document
model = sw._active_doc()

dt = SW_DesignTable(model, r'G:\My Drive\Google Drive - Work\TestMacros')

dt.update_from_csv()

sw_model = sw.get_model()
sw_model.extension.rebuild(1)
print(f"\n\nDone.\n\n")