import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import googlemaps
import uuid
import time
import pandas as pd
import os

class DataCrawlerApp:
    def __init__(self, master):
        self.master = master
        master.title("Google Maps Data Crawler")
        
        # Set the window icon
        master.iconbitmap('./res/company_icon.ico')  # Ensure the icon file is in the same directory or provide the correct path
        
        # Initialize a set to track already fetched Google Map Links
        self.existing_links = set()
        self.input_file_path = None
        self.gmaps = None  # Will be set after user enters API key

        # API Key Input
        tk.Label(master, text="API Key:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.api_key_entry = tk.Entry(master, width=50)
        self.api_key_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Search query input
        tk.Label(master, text="Search Query:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.query_entry = tk.Entry(master, width=50)
        self.query_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Number of records input
        tk.Label(master, text="Records to Fetch:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.num_records_var = tk.IntVar(value=10)
        self.num_records_spinbox = tk.Spinbox(master, from_=1, to=100, textvariable=self.num_records_var, width=5)
        self.num_records_spinbox.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Fetched count display
        tk.Label(master, text="Fetched Count:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.fetched_count_var = tk.StringVar(value="0")
        self.fetched_count_label = tk.Label(master, textvariable=self.fetched_count_var)
        self.fetched_count_label.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # Option for continuing fetching
        self.continue_fetching_var = tk.BooleanVar(value=True)
        self.continue_fetching_check = tk.Checkbutton(master, text="Continue Fetching if More Data Available", variable=self.continue_fetching_var)
        self.continue_fetching_check.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        # Field for selecting an existing input file
        tk.Label(master, text="Existing Data File:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.file_label = tk.Label(master, text="No file selected")
        self.file_label.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.load_file_button = tk.Button(master, text="Load File", command=self.load_input_file)
        self.load_file_button.grid(row=5, column=2, padx=5, pady=5)
        
        # Treeview widget to display fetched data
        self.tree = ttk.Treeview(master, columns=("Name", "Address", "Phone", "Website", "Rating", "Google Map Link"), show="headings")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Address", text="Address")
        self.tree.heading("Phone", text="Phone")
        self.tree.heading("Website", text="Website")
        self.tree.heading("Rating", text="Rating")
        self.tree.heading("Google Map Link", text="Google Map Link")
        self.tree.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        
        # Buttons for fetching data
        self.fetch_button = tk.Button(master, text="Fetch Data", command=self.start_fetching)
        self.fetch_button.grid(row=7, column=0, padx=5, pady=10)
        self.continue_button = tk.Button(master, text="Continue Fetching", command=self.continue_fetching, state=tk.DISABLED)
        self.continue_button.grid(row=7, column=1, padx=5, pady=10)
        
        # Container to store fetched data (list of dictionaries)
        self.fetched_data = []
        self.next_page_token = None
        self.is_fetching = False

        # Add "Developed by" section
        dev_label = tk.Label(master, text="Developed by ", fg="black")
        dev_label.grid(row=8, column=0, padx=5, pady=10, sticky="w")
        link_label = tk.Label(master, text="Bright Techno Solutions", fg="blue", cursor="hand2")
        link_label.grid(row=8, column=1, pady=10, sticky="w")
        tk.Label(master, text="© 2025").grid(row=8, column=2, pady=10, sticky="w")
        link_label.bind("<Button-1>", lambda e: os.system("start https://www.brighttechnosolutions.com"))

    def load_input_file(self):
        """Open a file dialog to select an existing Excel/CSV file and load previously fetched data."""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            try:
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                else:
                    df = pd.read_csv(file_path)
                # Assume that the 'Google map Link' field is used to identify a record uniquely
                if 'Google map Link' in df.columns:
                    self.existing_links = set(df['Google map Link'].dropna().tolist())
                messagebox.showinfo("File Loaded", f"Loaded {len(self.existing_links)} existing records.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def start_fetching(self):
        # Check API key
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            messagebox.showerror("Error", "Please enter your API key.")
            return
        
        # Create a new Google Maps client using the provided API key
        self.gmaps = googlemaps.Client(key=api_key)
        
        query = self.query_entry.get().strip()
        if not query:
            messagebox.showerror("Error", "Please enter a search query.")
            return
        
        # Reset interface and data (keep existing_links loaded from file)
        self.fetch_button.config(state=tk.DISABLED)
        self.continue_button.config(state=tk.DISABLED)
        self.fetched_data.clear()
        self.tree.delete(*self.tree.get_children())
        self.fetched_count_var.set("0")
        self.next_page_token = None
        self.is_fetching = True
        
        # Start data fetching in a new thread to keep the GUI responsive
        threading.Thread(target=self.fetch_data, args=(query, self.num_records_var.get()), daemon=True).start()

    def fetch_data(self, query, num_records):
        fetched_count = len(self.fetched_data)
        results = []
        try:
            # Initial search request using the Places API
            response = self.gmaps.places(query=query)
            results.extend(response.get("results", []))
            self.next_page_token = response.get("next_page_token", None)
        except Exception as e:
            messagebox.showerror("API Error", str(e))
            self.is_fetching = False
            self.fetch_button.config(state=tk.NORMAL)
            return
        
        # Loop until desired number of new (non-duplicate) records is reached or no more pages
        while fetched_count < num_records and results:
            for place in results:
                if fetched_count >= num_records:
                    break
                place_id = place.get("place_id")
                details = {}
                try:
                    details_response = self.gmaps.place(place_id=place_id)
                    details = details_response.get("result", {})
                except Exception as e:
                    print("Error fetching details:", e)
                name = details.get("name", "")
                address = details.get("formatted_address", "")
                phone = details.get("formatted_phone_number", "")
                website = details.get("website", "")
                rating = details.get("rating", "N/A")
                google_map_link = f"https://www.google.com/maps/place/?q=place_id:{place_id}"
                
                # Skip if already fetched
                if google_map_link in self.existing_links:
                    continue

                # Insert data into the treeview
                self.tree.insert("", "end", values=(name, address, phone, website, rating, google_map_link))
                fetched_count += 1
                record = {
                    "Lead ID": str(uuid.uuid4()),
                    "Name": name,
                    "Phone Number": phone,
                    "Phone Number 2": "",
                    "Email": "",
                    "Insitute Name": name,
                    "Address": address,
                    "Subscription Plan": "",
                    "Payment Status": "",
                    "Source Of Enqiry": "Google Maps",
                    "Status": "",
                    "Subscription Taken": "",
                    "Notes": "",
                    "Google map Link": google_map_link,
                    "Website": website,
                    "Website Review": "",
                    "Rating": rating
                }
                self.fetched_data.append(record)
                # Add to existing links set so duplicates aren’t fetched again
                self.existing_links.add(google_map_link)
                self.fetched_count_var.set(str(fetched_count))
                self.master.update_idletasks()
            
            # Check for additional pages if needed
            if fetched_count < num_records and self.next_page_token:
                time.sleep(2)  # Pause to allow the next_page_token to become active
                try:
                    response = self.gmaps.places(query=query, page_token=self.next_page_token)
                    results = response.get("results", [])
                    self.next_page_token = response.get("next_page_token", None)
                except Exception as e:
                    print("Error during pagination:", e)
                    break
            else:
                break
        
        self.is_fetching = False
        self.fetch_button.config(state=tk.NORMAL)
        # Enable continue button if more data is available and the option is checked
        if self.next_page_token and self.continue_fetching_var.get():
            self.continue_button.config(state=tk.NORMAL)
        
        # After fetching, if an input file is loaded, append the new data to it.
        if self.input_file_path and self.fetched_data:
            self.append_to_file()

    def continue_fetching(self):
        if not self.is_fetching and self.next_page_token:
            self.fetch_button.config(state=tk.DISABLED)
            self.continue_button.config(state=tk.DISABLED)
            query = self.query_entry.get().strip()
            threading.Thread(target=self.fetch_data, args=(query, self.num_records_var.get()), daemon=True).start()

    def append_to_file(self):
        """Append the newly fetched data to the loaded input file."""
        try:
            if self.input_file_path.lower().endswith(('.xlsx', '.xls')):
                try:
                    df_existing = pd.read_excel(self.input_file_path)
                except Exception:
                    df_existing = pd.DataFrame()
            else:
                try:
                    df_existing = pd.read_csv(self.input_file_path)
                except Exception:
                    df_existing = pd.DataFrame()
                    
            # Create a DataFrame from the newly fetched data
            df_new = pd.DataFrame(self.fetched_data)
            # Combine with existing data
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            # Write back to the same file
            if self.input_file_path.lower().endswith(('.xlsx', '.xls')):
                df_combined.to_excel(self.input_file_path, index=False)
            else:
                df_combined.to_csv(self.input_file_path, index=False)
            messagebox.showinfo("File Updated", f"New data appended to {os.path.basename(self.input_file_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to append data to file:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    # Optional: Make the GUI responsive by configuring grid weights
    root.grid_rowconfigure(6, weight=1)
    root.grid_columnconfigure(1, weight=1)
    app = DataCrawlerApp(root)
    root.mainloop()
