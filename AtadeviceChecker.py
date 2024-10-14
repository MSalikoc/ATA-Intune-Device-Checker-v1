import tkinter as tk
from tkinter import messagebox, ttk
import msal
import requests
import csv
import threading
import os

# Global variables
DEVICE_DATA = []
GRAPH_API_URL = 'https://graph.microsoft.com/v1.0/'
selected_search_value = ""
search_type = "deviceName"  # Default search type

# Function to authenticate with Microsoft Graph
def authenticate(client_id, tenant_id):
    authority = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.PublicClientApplication(client_id, authority=authority)

    scopes = ["Device.Read.All", "DeviceManagementManagedDevices.Read.All"]
    flow = app.initiate_device_flow(scopes=scopes)

    if "user_code" not in flow:
        messagebox.showerror("Error", "Failed to create device flow")
        return None

    web_auth_url = flow['verification_uri']
    web_auth_button = messagebox.askyesno("Open Browser", "Do you want to open the browser for authentication?")
    
    if web_auth_button:
        import webbrowser
        webbrowser.open(web_auth_url)
    
    messagebox.showinfo("Authenticate", f"Please authenticate using this code: {flow['user_code']}")
    
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        return result["access_token"]
    else:
        messagebox.showerror("Error", f"Authentication failed: {result.get('error_description', 'Unknown error')}")
        return None

# Fetch device details with paging support
def fetch_device_details(token):
    global DEVICE_DATA
    DEVICE_DATA.clear()  # Clear previous data
    headers = {'Authorization': f'Bearer {token}'}
    url = GRAPH_API_URL + 'deviceManagement/managedDevices'
    
    while url:
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            DEVICE_DATA.extend(data.get('value', []))
            url = data.get('@odata.nextLink', None)  # Check if there is another page
        else:
            messagebox.showerror("Error", f"Failed to fetch devices: {response.text}")
            return False
    
    messagebox.showinfo("Success", "Device data fetched successfully")
    return True

# Class to manage pages
class Page(tk.Frame):
    def __init__(self, *args, **kwargs):
        tk.Frame.__init__(self, *args, **kwargs)
    def show(self):
        self.lift()

# Authentication Page with Client ID and Tenant ID Input
class AuthPage(Page):
    def __init__(self, *args, **kwargs):
        Page.__init__(self, *args, **kwargs)
        label = tk.Label(self, text="Intune Device Checker Authentication Page", font=("Arial", 18), pady=5)
        label.pack(side="top", fill="both", expand=True)

        label = tk.Label(self, text="© Ali Koc", font=("Arial", 12), pady=5)
        label.pack(side="top", fill="both", expand=True)

        # Input fields for Client ID and Tenant ID
        self.client_id_label = tk.Label(self, text="Client ID:", pady=5)
        self.client_id_label.pack()
        self.client_id_entry = tk.Entry(self)
        self.client_id_entry.pack()

        self.tenant_id_label = tk.Label(self, text="Tenant ID:", pady=5)
        self.tenant_id_label.pack()
        self.tenant_id_entry = tk.Entry(self)
        self.tenant_id_entry.pack()

        self.auth_button = tk.Button(self, text="Authenticate", command=self.authenticate, pady=5, padx=20)
        self.auth_button.pack(pady=10)

    def authenticate(self):
        client_id = self.client_id_entry.get()
        tenant_id = self.tenant_id_entry.get()
        if not client_id or not tenant_id:
            messagebox.showerror("Error", "Client ID and Tenant ID are required!")
            return

        token = authenticate(client_id, tenant_id)
        if token:
            self.master.token = token
            messagebox.showinfo("Success", "Authentication Successful")
            self.master.device_page.show()

# Device Name or User Input Page
class DeviceInputPage(Page):
    def __init__(self, *args, **kwargs):
        Page.__init__(self, *args, **kwargs)
        label = tk.Label(self, text="Search Page", font=("Arial", 18), pady=10)
        label.pack(side="top", fill="both", expand=True)

        self.search_input = tk.Entry(self)
        self.search_input.pack(pady=10)

        self.search_type = tk.StringVar(value="deviceName")
        tk.Radiobutton(self, text="Device Name", variable=self.search_type, value="deviceName").pack()
        tk.Radiobutton(self, text="User Name", variable=self.search_type, value="userPrincipalName").pack()

        self.search_button = tk.Button(self, text="Search", command=self.search_device, pady=5, padx=10)
        self.search_button.pack(pady=10)

        self.back_button = tk.Button(self, text="Back", command=lambda: self.master.auth_page.show(), pady=5, padx=10)
        self.back_button.pack(pady=10)

    def search_device(self):
        global selected_search_value, search_type
        selected_search_value = self.search_input.get()
        search_type = self.search_type.get()
        
        if selected_search_value:
            messagebox.showinfo("Success", f"Searching by {search_type}: {selected_search_value}")
            self.master.data_page.show()
        else:
            messagebox.showerror("Error", "Please enter a valid search value")

# Data Fetch Page with threading and design updates
class DataPage(Page):
    def __init__(self, *args, **kwargs):
        Page.__init__(self, *args, **kwargs)
        label = tk.Label(self, text="Device Information Check", font=("Arial", 18), pady=10)
        label.pack(side="top", fill="both", expand=True)

        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg="lightgrey")
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar_x = tk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        scrollbar_y = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollbar_x.pack(side="bottom", fill="x")
        scrollbar_y.pack(side="right", fill="y")

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background="lightblue")
        style.configure("Treeview", font=("Arial", 10), background="white", foreground="black", rowheight=25)

        # Updated column order as requested
        self.tree = ttk.Treeview(canvas, columns=("Device Name", "Model", "Compliance Status", "Last Sync Date Time", 
                                                  "Device Manufacturer", "Operating System", "OS Version", 
                                                  "Serial Number", "Ownership"), show="headings")

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        canvas.create_window((0, 0), window=self.tree, anchor="nw")
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"), xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        self.btn_fetch = tk.Button(self, text="Fetch Devices", command=self.fetch_device_details_async, pady=5, padx=10)
        self.btn_fetch.pack(pady=10)

        self.btn_export = tk.Button(self, text="Export to CSV", command=self.export_to_csv, pady=5, padx=10)
        self.btn_export.pack(pady=10)

        self.back_button = tk.Button(self, text="Back", command=lambda: self.master.device_page.show(), pady=5, padx=10)
        self.back_button.pack(pady=10)

    def fetch_device_details_async(self):
        thread = threading.Thread(target=self.fetch_device_details)
        thread.start()

    def fetch_device_details(self):
        token = self.master.token
        if not token:
            messagebox.showerror("Error", "Please authenticate first")
            return

        success = fetch_device_details(token)
        if success:
            self.display_data()

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        found_device = False
        for i, device in enumerate(DEVICE_DATA):
            if device.get(search_type, "").lower() == selected_search_value.lower():
                found_device = True
                values = [
                    device.get("deviceName", 'N/A'),
                    device.get("model", 'N/A'),
                    device.get("complianceState", 'N/A'),
                    device.get("lastSyncDateTime", 'N/A'),
                    device.get("manufacturer", 'N/A'),
                    device.get("operatingSystem", 'N/A'),
                    device.get("osVersion", 'N/A'),
                    device.get("serialNumber", 'N/A'),
                    device.get("ownership", 'N/A')
                ]
                
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                self.tree.insert("", tk.END, values=values, tags=(tag,))
        
        if not found_device:
            messagebox.showinfo("Not Found", f"No data found for {search_type}: {selected_search_value}")
        
        self.tree.tag_configure("evenrow", background="lightgrey")
        self.tree.tag_configure("oddrow", background="white")

    def export_to_csv(self):
        file_path = os.path.join("C:\\", "intune_device_data.csv")
        with open(file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Device Name", "Model", "Compliance Status", "Last Sync Date Time", 
                             "Device Manufacturer", "Operating System", "OS Version", 
                             "Serial Number", "Ownership"])
            for device in DEVICE_DATA:
                if device.get(search_type, "").lower() == selected_search_value.lower():
                    writer.writerow([
                        device.get("deviceName", ""),
                        device.get("model", ""),
                        device.get("complianceState", ""),
                        device.get("lastSyncDateTime", ""),
                        device.get("manufacturer", ""),
                        device.get("operatingSystem", ""),
                        device.get("osVersion", ""),
                        device.get("serialNumber", ""),
                        device.get("ownership", "")
                    ])
        messagebox.showinfo("Export", f"Data exported to {file_path}")

# Main Application
class MainApplication(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        tk.Frame.__init__(self, master, *args, **kwargs)
        self.token = None
        self.pack(fill="both", expand=True)

        # Create pages
        self.auth_page = AuthPage(self)
        self.device_page = DeviceInputPage(self)
        self.data_page = DataPage(self)

        # Add copyright label to the top-right corner
        self.copyright_label = tk.Label(self, text="Copyright Ali Koc", font=("Arial", 8))
        self.place_widget_relative(self.copyright_label)

        # Stack pages
        self.auth_page.place(in_=self, x=0, y=0, relwidth=1, relheight=1)
        self.device_page.place(in_=self, x=0, y=0, relwidth=1, relheight=1)
        self.data_page.place(in_=self, x=0, y=0, relwidth=1, relheight=1)

        self.auth_page.show()

    def place_widget_relative(self, widget):
        widget.place(relx=0.98, rely=0.02, anchor="ne")  # Adjusted the position slightly

# Initialize the application
root = tk.Tk()
root.geometry("600x600")  # Set the default window size to be bigger for more data
root.title("ATA Microsoft Intune Device Checker")  # Updated application title
app = MainApplication(root)
root.mainloop()
