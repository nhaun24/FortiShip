import tkinter as tk
from tkinter import messagebox, ttk
import configparser
import csv
import os
from datetime import datetime
import json
import re
import time
import logging
import subprocess
from pathlib import Path

# Set up logging
logging.basicConfig(
    filename='label_printer.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class PrinterError(Exception):
    pass

class DeviceLogger:
    def __init__(self, filename):
        self.filename = filename
        self.ensure_csv_exists()
        self.backup_dir = Path("backups")
        self.backup_dir.mkdir(exist_ok=True)
    
    def ensure_csv_exists(self):
        if not os.path.exists(self.filename):
            with open(self.filename, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Timestamp', 'Site', 'AP Label', 'Serial Number', 'MAC Address'])
    
    def log_device(self, site, ap_label, serial_number, mac=''):
        with open(self.filename, 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow([
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                site,
                ap_label,
                serial_number,
                mac
            ])
        logging.info(f"Logged device: {ap_label} - {serial_number}")
        self._create_backup()

    def _create_backup(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = self.backup_dir / f"device_log_backup_{timestamp}.csv"
        try:
            with open(self.filename, 'r') as source, open(backup_file, 'w') as target:
                target.write(source.read())
            logging.info(f"Created backup: {backup_file}")
        except Exception as e:
            logging.error(f"Backup failed: {e}")

class LabelPrinter:
    def __init__(self):
        self.ptexe_path = r"C:\Program Files (x86)\Brother\Ptedit54\ptedit54.exe"
        self.template_path = os.path.join(os.path.dirname(__file__), "template.lbx")
        
        if not os.path.exists(self.ptexe_path):
            alternative_paths = [
                r"C:\Program Files (x86)\Brother\P-touch Editor 5.4\PtCmd.exe",
            ]
            for path in alternative_paths:
                if os.path.exists(path):
                    self.ptexe_path = path
                    break
        
        logging.info(f"Using printer executable: {self.ptexe_path}")

    def print_label(self, ap_label, mac, serial):
      try:
          if not os.path.exists(self.ptexe_path):
              raise PrinterError(f"Brother P-touch Editor not found at {self.ptexe_path}")

          # Command to print using PtCmd
          cmd = [
              self.ptexe_path,
              "/ff",          # Form feed
              "/c",           # Close after printing
              "/d",          # Use default printer
              self.template_path,
              "/v",          # Set variables
              f"ap_label={ap_label}",
              f"mac={mac}",
              f"serial={serial}",
              "/p"           # Print command
          ]

          logging.info(f"Executing print command: {' '.join(cmd)}")
          result = subprocess.run(cmd, 
                               capture_output=True, 
                               text=True, 
                               creationflags=subprocess.CREATE_NO_WINDOW)

          if result.stdout:
              logging.info(f"Print output: {result.stdout}")
          if result.stderr:
              logging.error(f"Print error: {result.stderr}")

          return result.returncode == 0

      except Exception as e:
          logging.error(f"Print error: {e}")
          return False

    def print_switch_label(self, ap_label, serial):
      try:
          if not os.path.exists(self.ptexe_path):
              raise PrinterError(f"Brother P-touch Editor not found at {self.ptexe_path}")

          # Command to print using PtCmd
          cmd = [
              self.ptexe_path,
              "/ff",          # Form feed
              "/c",           # Close after printing
              "/d",          # Use default printer
              "switch_template.lbx",
              "/v",          # Set variables
              f"ap_label={ap_label}",
              f"mac=",
              f"serial={serial}",
              "/p"           # Print command
          ]

          logging.info(f"Executing print command: {' '.join(cmd)}")
          result = subprocess.run(cmd, 
                               capture_output=True, 
                               text=True, 
                               creationflags=subprocess.CREATE_NO_WINDOW)

          if result.stdout:
              logging.info(f"Print output: {result.stdout}")
          if result.stderr:
              logging.error(f"Print error: {result.stderr}")

          return result.returncode == 0

      except Exception as e:
          logging.error(f"Print error: {e}")
          return False

    def test_print(self):
        return self.print_label("TEST-AP", "00:11:22:33:44:55", "TEST123")

    def test_print(self):
        return self.print_label("TEST-AP", "00:11:22:33:44:55", "TEST123")

class LabelPrinterTool(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Label Printer Tool")
        self.config(bg='#2578be')
        
        # Initialize counters
        self.ap_counter = 1
        self.batch_entries = []
        self.switch_counter = "A"
        
        # Initialize printer and logger
        self.printer = LabelPrinter()
        self.logger = DeviceLogger('device_log.csv')
        
        # Load sites
        self.load_sites()
        
        # Build UI
        self.setup_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        # Center window
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')

    def load_sites(self):
        try:
            with open('sites.json', 'r') as f:
                self.sites = json.load(f)
        except Exception as e:
            logging.error(f"Failed to load sites: {e}")
            self.sites = [{"name": "Default Site"}]

    def setup_ui(self):
        # Main Frame
        main_frame = tk.Frame(self, bg='#2578be')
        main_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)
        
        # Header
        tk.Label(main_frame, text="Fortinet Label & Batch Tool", 
                bg='#2578be', fg='#edf2ff',
                font=("Comic Sans MS", 20, "bold")).grid(row=0, column=0, 
                columnspan=2, pady=20)

        # Site Selection
        tk.Label(main_frame, text="Select Site:", bg='#2578be', font=("Ariel", 12),
                fg='#edf2ff').grid(row=1, column=0, sticky='e')
        self.site_dropdown = ttk.Combobox(main_frame, 
                                        values=[s['name'] for s in self.sites],
                                        state='readonly')
        self.site_dropdown.grid(row=1, column=1, sticky='ew', pady=5)
        self.site_dropdown.bind('<<ComboboxSelected>>', self.on_site_select)

        # Input Fields
        tk.Label(main_frame, text="Serial Number:", bg='#2578be', font=("Ariel", 12),
                fg='#edf2ff').grid(row=2, column=0, sticky='e')
        self.serial_number_entry = tk.Entry(main_frame, bg='#edf2ff', fg='#000000')
        self.serial_number_entry.grid(row=2, column=1, sticky='ew', pady=5)
        self.serial_number_entry.bind('<Return>', self.handle_serial_enter)

        tk.Label(main_frame, text="MAC Address:", bg='#2578be', font=("Ariel", 12),
                fg='#edf2ff').grid(row=3, column=0, sticky='e')
        self.mac_entry = tk.Entry(main_frame, bg='#edf2ff', fg='#000000')
        self.mac_entry.grid(row=3, column=1, sticky='ew', pady=5)
        self.mac_entry.bind('<Return>', lambda e: self.add_to_batch())

        #tk.Label(main_frame, text="Switch Serial:", bg='#2578be', 
        #        fg='#edf2ff').grid(row=4, column=0, sticky='e')
        #self.switch_entry = tk.Entry(main_frame, bg='#616364', fg='#edf2ff')
        #self.switch_entry.grid(row=4, column=1, sticky='ew', pady=5)
        #self.switch_entry.bind('<Return>', lambda e: self.add_to_batch())

        # Batch List
        self.ap_listbox = tk.Listbox(main_frame, height=10, width=70, 
                                    bg='#edf2ff', fg='#000000')
        self.ap_listbox.grid(row=5, column=0, columnspan=2, pady=10,sticky="nsew")

        # Buttons Frame
        button_frame = tk.Frame(main_frame, bg='#2578be')
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)

        #tk.Button(button_frame, text="Add to Batch", 
        #         command=self.add_to_batch,
        #         bg='#3fe010', fg='#edf2ff').pack(side=tk.LEFT, padx=5)
                 
        tk.Button(button_frame, text="Export CSV", 
                 command=self.create_csv,
                 bg='#0c8714', fg='#edf2ff').pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="Test Print",
         command=self.test_print,
         bg='#b06504', fg='#edf2ff').pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="Remove Last", 
                 command=self.remove_last_entry,
                 bg='#8a7a06', fg='#edf2ff').pack(side=tk.LEFT, padx=5)
                 
        tk.Button(button_frame, text="Clear All", 
                 command=self.reset_counters,
                 bg='#e01b10', fg='#edf2ff').pack(side=tk.LEFT, padx=5)
                 

        #tk.Button(button_frame, text="Debug Printer",
        #         command=self.debug_printer,
        #         bg='#4287f5', fg='#edf2ff').pack(side=tk.LEFT, padx=5)

        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = tk.Label(main_frame, textvariable=self.status_var,
                                 bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=7, column=0, columnspan=2, sticky='ew')

        #self.bind("<Configure>", self.resize_font)

        # Configure grid weights
        main_frame.grid_rowconfigure(0, weight=1)  # header row
        main_frame.grid_rowconfigure(1, weight=0)  # site selection
        main_frame.grid_rowconfigure(2, weight=0)  # serial number
        main_frame.grid_rowconfigure(3, weight=0)  # mac address
        main_frame.grid_rowconfigure(4, weight=0)  # (you can add more rows as needed)
        main_frame.grid_rowconfigure(5, weight=1)  # ap listbox
        main_frame.grid_columnconfigure(0, weight=1)  # Left column (labels)
        main_frame.grid_columnconfigure(1, weight=2)  # Right column (inputs)

    def on_site_select(self, event):
        self.serial_number_entry.focus()

    def handle_serial_enter(self, event):
        serial = self.serial_number_entry.get().strip()
        if serial.upper().startswith("FP"):
            self.mac_entry.focus()
        else:
            self.add_to_batch()

    def add_to_batch(self):
        try:
            site = self.site_dropdown.get()
            serial = self.serial_number_entry.get().strip()
            mac = self.mac_entry.get().strip()

            if not site:
                messagebox.showwarning("Input Error", "Please select a site")
                return

            if not serial:
                messagebox.showwarning("Input Error", "Serial number is required")
                return

            if serial.upper().startswith("S") and not mac:
                # Switch with no MAC is OK
                pass
            elif serial.upper().startswith("FP") and not mac:
                # AP with missing MAC = ERROR
                messagebox.showwarning("Input Error", "MAC address is required for APs")
                return
            elif not serial.upper().startswith(("FP", "S")):
                messagebox.showwarning("Input Error", "Invalid serial number")
                self.serial_number_entry.delete(0, tk.END)
                self.mac_entry.delete(0, tk.END)
                self.serial_number_entry.focus()
                return

            if not self.is_unique_device(serial, mac):
                messagebox.showerror("Duplicate Error", "Device has already been scanned")
                self.serial_number_entry.delete(0, tk.END)
                self.mac_entry.delete(0, tk.END)
                self.serial_number_entry.focus()
                return

            if serial.upper().startswith("FP"):
                # It's an AP
                ap_label = f"AP-{format(self.ap_counter, '02')}"

                # Create or overwrite print.csv with single entry
                with open('print.csv', 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['AP Label', 'Serial Number', 'MAC'])
                    writer.writerow([ap_label, serial, mac])

                logging.info(f"Attempting to print AP label: {ap_label}")
                self.printer.print_label(ap_label, mac, serial)

                # Log and list
                entry = f"{ap_label} - Serial Number: {serial} - MAC: {mac}"
                self.logger.log_device(site, ap_label, serial, mac)
                self.ap_listbox.insert(tk.END, entry)

                # Increment AP counter
                self.ap_counter += 1

            elif serial.upper().startswith("S"):
                # It's a Switch
                switch_label = f"{self.switch_counter}"

                # Create or overwrite print.csv with single entry
                with open('switch_print.csv', 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Switch Label', 'Serial Number'])
                    writer.writerow([switch_label, serial])

                logging.info(f"Attempting to print Switch label: {switch_label}")
                self.printer.print_switch_label(switch_label, serial)

                # Log and list
                entry = f"{switch_label} - Serial Number: {serial}"
                self.logger.log_device(site, switch_label, serial, mac)
                self.ap_listbox.insert(tk.END, entry)

                # Increment Switch counter
                self.switch_counter = self.increment_letter(self.switch_counter)

            # Clear entries
            self.serial_number_entry.delete(0, tk.END)
            self.mac_entry.delete(0, tk.END)
            self.serial_number_entry.focus()

            self.status_var.set("Label printed successfully")

        except Exception as e:
            logging.error(f"Error in add_to_batch: {e}")
            messagebox.showerror("Error", str(e))
            self.status_var.set(f"Error: {str(e)}")

    def is_unique_device(self, serial, mac):
        try:
            with open('device_log.csv', 'r', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get('Serial Number') == serial or row.get('MAC') == mac:
                        return False
            return True
        except FileNotFoundError:
            # If the log doesn't exist yet, we assume it's unique
            return True
        except Exception as e:
            logging.error(f"Error checking for duplicate device: {e}")
            raise

    def increment_letter(self, letter):
        if letter.upper() == "Z":
            return "A"  # or throw an error if you don't want it to wrap
        return chr(ord(letter.upper()) + 1)

    def test_print(self):
        if self.printer.test_print():
            self.status_var.set("Test print successful")
            messagebox.showinfo("Success", "Test print completed successfully")
        else:
            self.status_var.set("Test print failed")
            messagebox.showerror("Error", "Test print failed. Check printer connection and logs")

    def debug_printer(self):
        debug_info = []
        debug_info.append(f"P-touch Editor path: {self.printer.ptexe_path}")
        debug_info.append(f"Path exists: {os.path.exists(self.printer.ptexe_path)}")
        debug_info.append(f"\nTemplate path: {self.printer.template_path}")
        debug_info.append(f"Template exists: {os.path.exists(self.printer.template_path)}")
        
        messagebox.showinfo("Printer Debug Info", "\n".join(debug_info))

    def remove_last_entry(self):
        if self.ap_listbox.size() > 0:
            last_entry = self.ap_listbox.get(tk.END)
    
            # Check what the last entry was (AP or Switch)
            if last_entry.startswith("AP-"):
                if self.ap_counter > 1:
                    self.ap_counter -= 1
            elif last_entry.startswith("SW-"):
                if self.switch_counter > "A":
                    self.switch_counter = self.decrement_letter(self.switch_counter)
    
            self.ap_listbox.delete(tk.END)
            self.status_var.set("Last entry removed")
        else:
            messagebox.showwarning("No Entries", "There are no entries to remove.")

    def reset_counters(self):
        if messagebox.askyesno("Clear All", "Are you sure you want to clear all entries?"):
            self.ap_counter = 1
            self.switch_counter = "A"
            self.ap_listbox.delete(0, tk.END)
            self.status_var.set("All entries cleared")

    def create_csv(self):
        selected_site = self.site_dropdown.get()

        if not selected_site:
            messagebox.showwarning("Missing Site", "Please select a site first.")
            return

        #Ensure the 'CSV' subfolder exists
        base_dir = os.path.dirname(os.path.abspath(__file__))  # Folder where script is
        csv_folder = os.path.join(base_dir, "CSV")
        os.makedirs(csv_folder, exist_ok=True)

        # Filename with timestamp (optional)
        #timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S"
        save_filename = os.path.join(csv_folder,f"/{selected_site}.csv")

        try:
            with open(save_filename, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(['Timestamp', 'Site', 'Device Label', 'Serial Number', 'MAC Address'])

                for index in range(self.ap_listbox.size()):
                    entry = self.ap_listbox.get(index)

                    if entry.startswith("AP-"):
                        # AP format: AP-XX - Serial Number: <serial> - MAC: <mac>
                        ap_label, serial_number_mac = entry.split(" - Serial Number: ")
                        serial_number, mac = serial_number_mac.split(" - MAC: ")
                    elif entry.startswith("SW-"):
                        # Switch format: SW-X - Serial Number: <serial> 
                        ap_label, serial_number = entry.split(" - Serial Number: ")
                        mac = ""  # Switches have no MAC
                    else:
                        continue  # Unknown format, skip

                    writer.writerow([
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        selected_site,
                        ap_label,
                        serial_number,
                        mac
                    ])

            messagebox.showinfo("CSV Export", f"Data exported to {save_filename}")
            self.status_var.set(f"Exported to {save_filename}")

        except Exception as e:
            messagebox.showerror("CSV Error", f"Failed to create CSV: {e}")

    def on_close(self):
        if messagebox.askokcancel("Quit", "Please ensure you have Exported to Batch before quiting."):
            #self.create_csv()
            self.destroy()

def main():
    app = LabelPrinterTool()
    app.mainloop()

if __name__ == "__main__":
    main()
