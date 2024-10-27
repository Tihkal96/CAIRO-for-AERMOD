import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Tk
import os
import subprocess
import shutil
import webbrowser
import win32clipboard
import time
import threading
import utm
import simplekml
from shutil import copyfile
import json
import sys


def load_config():
    config_path = os.path.join(os.getcwd(), 'config.json')
    try:
        with open(config_path, 'r') as config_file:
            config = json.load(config_file)
            return config
    except FileNotFoundError:
        print("Configuration file not found. Please make sure 'config.json' is present in the project folder.")
        return None


config = load_config()
if config is None:
    sys.exit()  # Exit if configuration cannot be loaded


def app1():
    def browse_files():
        filename = filedialog.askopenfilename()
        if filename:
            file_name = os.path.basename(filename)
            new_entry = ttk.Entry(datafile_frame)
            new_entry.grid(row=len(datafile_entries), column=1, sticky="we")
            new_entry.insert(0, file_name)
            datafile_entries.append((new_entry, filename))

    def generate_output():
        output_text_content = ""
        output_text_content += "CO STARTING\n"

        # Check if title_entry is not empty
        if title_entry.get():
            output_text_content += "CO TITLEONE  " + title_entry.get() + "\n"

        # Check if datafile_entries is not empty
        if datafile_entries:
            output_text_content += "CO DATATYPE  " + datatype_combobox.get()
            if datatype_combobox.get() == "DEM":
                output_text_content += "     FILLGAPS\n"
            else:
                output_text_content += "\n"

            for entry, full_path in datafile_entries:
                filename = os.path.basename(full_path)
                output_text_content += "CO DATAFILE  " + filename + "\n"

        if (anchor_lat_entry.get() and
                anchor_long_entry.get() and
                utm_zone_entry.get() and
                utm_datum_entry.get()):
            output_text_content += ("   ANCHORXY  " +
                                    anchor_long_entry.get() + " " +
                                    anchor_lat_entry.get() + " " +
                                    anchor_long_entry.get() + " " +
                                    anchor_lat_entry.get() + " " +
                                    utm_zone_entry.get() + " " +
                                    utm_datum_entry.get() + "\n")

        if flagpole_entry.get():
            output_text_content += "   FLAGPOLE  " + flagpole_entry.get() + "\n"

        output_text_content += "CO RUNORNOT  RUN\n"
        output_text_content += "CO FINISHED\n"
        output_text_content += "RE STARTING\n"

        if (anchor_lat_entry.get() and
                anchor_long_entry.get() and
                x_spacing_entry.get() and
                x_length_entry.get() and
                y_spacing_entry.get() and
                y_length_entry.get()):
            output_text_content += "   GRIDCART CART01 STA\n"
            output_text_content += ("                    XYINC " +
                                    anchor_long_entry.get() + " " +
                                    x_spacing_entry.get() + " " +
                                    x_length_entry.get() + " " +
                                    anchor_lat_entry.get() + " " +
                                    y_spacing_entry.get() + " " +
                                    y_length_entry.get() + "\n")
            output_text_content += "   GRIDCART CART01 END\n"

        output_text_content += "RE FINISHED\n"
        output_text_content += "OU STARTING\n"
        output_text_content += "   RECEPTOR  RECEPT.ROU\n"
        output_text_content += "OU FINISHED\n"

        return output_text_content

    def compile_output():
        output_text_content = generate_output()
        folder_path = filedialog.askdirectory()
        if folder_path:
            file_path = os.path.join(folder_path, "aermap.inp")
            with open(file_path, "w") as file:
                file.write(output_text_content)
            for entry, full_path in datafile_entries:
                _, file_name = os.path.split(full_path)
                destination = os.path.join(folder_path, file_name)
                if not os.path.exists(destination):
                    copyfile(full_path, destination)

            root.destroy()

    class Tooltip:
        def __init__(self, widget, text):
            self.widget = widget
            self.text = text
            self.tooltip = None
            self.widget.bind("<Enter>", self.enter)
            self.widget.bind("<Leave>", self.leave)

        def enter(self, event=None):
            x, y, _, _ = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 25
            if event:
                x = event.x_root + 10
                y = event.y_root + 10
            self.tooltip = tk.Toplevel(self.widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()

        def leave(self, event=None):
            if self.tooltip:
                self.tooltip.destroy()

    def get_clipboard_text():
        try:
            win32clipboard.OpenClipboard()
            clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
            win32clipboard.CloseClipboard()
            return clipboard_data.decode('utf-8')
        except (UnicodeDecodeError, TypeError):
            return ""

    def open_google_maps_for_anchor(anchor_lat_entry, anchor_long_entry):
        url = "https://www.google.com/maps"
        webbrowser.open(url)

        def monitor_clipboard():
            last_clipboard_text = get_clipboard_text()
            while True:
                clipboard_text = get_clipboard_text()
                if clipboard_text != last_clipboard_text:
                    last_clipboard_text = clipboard_text
                    try:
                        lat, lon = map(float, clipboard_text.split(','))
                        utm_coords = utm.from_latlon(lat, lon)
                        utm_easting, utm_northing, utm_zone_number, utm_zone_letter = utm_coords
                        if anchor_lat_entry.get() == '' and anchor_long_entry.get() == '':
                            anchor_lat_entry.delete(0, 'end')
                            anchor_long_entry.delete(0, 'end')
                            utm_zone_entry.delete(0, 'end')
                            utm_northing_rounded = round(utm_northing, 4)
                            utm_easting_rounded = round(utm_easting, 4)
                            anchor_lat_entry.insert(0, utm_northing_rounded)
                            anchor_long_entry.insert(0, utm_easting_rounded)
                            utm_zone_entry.insert(0, utm_zone_number)
                    except ValueError:
                        print("Invalid coordinates format in clipboard")
                time.sleep(1)

        clipboard_thread = threading.Thread(target=monitor_clipboard)
        clipboard_thread.daemon = True
        clipboard_thread.start()

    # Create the main window
    root = tk.Tk()
    root.title("AERMAP Input File Generator")

    # Create and pack the title entry
    title_label = ttk.Label(root, text="Title:", font=('Arial', 8))
    title_label.grid(row=0, column=0, sticky="e")
    title_entry = ttk.Entry(root, font=('Arial', 8))
    title_entry.grid(row=0, column=1, sticky="we")

    # Create and pack the datatype combobox
    datatype_label = ttk.Label(root, text="Data Type:", font=('Arial', 8))
    datatype_label.grid(row=1, column=0, sticky="e")
    Tooltip(datatype_label, "NED includes tiff files")
    datatype_combobox = ttk.Combobox(root, values=["NED", "DEM"], state="readonly", font=('Arial', 8))
    datatype_combobox.current(0)
    datatype_combobox.grid(row=1, column=1, sticky="we")

    # Create and pack the datafile entries
    datafile_entries = []
    datafile_label = ttk.Label(root, font=('Arial', 8))
    datafile_label.grid(row=2, column=0, sticky="e")
    datafile_frame = ttk.Frame(root)
    datafile_frame.grid(row=2, column=1, sticky="we")
    add_datafile_button = ttk.Button(datafile_frame, text="Add Data File",
                                     command=lambda: datafile_entries.append(ttk.Entry(datafile_frame)).grid(
                                         row=len(datafile_entries), column=0, sticky="we"))
    for entry in datafile_entries:
        entry.grid(row=datafile_entries.index(entry), column=1, sticky="we")

    # Create and pack the anchor coordinates entries
    choose_on_map_button = ttk.Button(root, text="Open map",
                                      command=lambda: open_google_maps_for_anchor(anchor_lat_entry, anchor_long_entry))
    choose_on_map_button.grid(row=3, column=1, sticky="we")
    Tooltip(choose_on_map_button, "Automatically inputs UTM zone and coordinates upon copying from Google maps;"
                                  " Specifies coordinates of bottom left grid corner")
    anchor_lat_label = ttk.Label(root, text="Anchor Northing:", font=('Arial', 8))
    anchor_lat_label.grid(row=4, column=0, sticky="e")
    Tooltip(anchor_lat_label, "Specifies Y coordinate of bottom left grid corner")
    anchor_lat_entry = ttk.Entry(root)
    anchor_lat_entry.grid(row=4, column=1, sticky="we")
    anchor_long_label = ttk.Label(root, text="Anchor Easting:", font=('Arial', 8))
    anchor_long_label.grid(row=5, column=0, sticky="e")
    Tooltip(anchor_long_label, "Specifies X coordinate of bottom left grid corner")
    anchor_long_entry = ttk.Entry(root)
    anchor_long_entry.grid(row=5, column=1, sticky="we")

    # Create and pack the UTM zone and datum entries
    utm_zone_label = ttk.Label(root, text="UTM Zone (0-60):", font=('Arial', 8))
    utm_zone_label.grid(row=6, column=0, sticky="e")
    utm_zone_entry = ttk.Entry(root)
    utm_zone_entry.grid(row=6, column=1, sticky="we")
    utm_datum_label = ttk.Label(root, text="UTM Datum:", font=('Arial', 8))
    utm_datum_label.grid(row=7, column=0, sticky="e")
    Tooltip(utm_datum_label, "0=No conversion")
    utm_datum_entry = ttk.Entry(root)
    utm_datum_entry.grid(row=7, column=1, sticky="we")

    # Create and pack the flagpole entry
    flagpole_label = ttk.Label(root, text="Flagpole Height (m):", font=('Arial', 8))
    flagpole_label.grid(row=8, column=0, sticky="e")
    flagpole_entry = ttk.Entry(root)
    flagpole_entry.grid(row=8, column=1, sticky="we")

    # Create and pack the x and y grid spacing and length entries
    x_spacing_label = ttk.Label(root, text="N°of X Gripoints:", font=('Arial', 8))
    x_spacing_label.grid(row=9, column=0, sticky="e")
    x_spacing_entry = ttk.Entry(root)
    x_spacing_entry.grid(row=9, column=1, sticky="we")
    x_length_label = ttk.Label(root, text="Grid \u0394X  (m):", font=('Arial', 8))
    x_length_label.grid(row=10, column=0, sticky="e")
    x_length_entry = ttk.Entry(root)
    x_length_entry.grid(row=10, column=1, sticky="we")
    y_spacing_label = ttk.Label(root, text="N°of Y Gridpoints:", font=('Arial', 8))
    y_spacing_label.grid(row=11, column=0, sticky="e")
    y_spacing_entry = ttk.Entry(root)
    y_spacing_entry.grid(row=11, column=1, sticky="we")
    y_length_label = ttk.Label(root, text="Grid \u0394Y (m):", font=('Arial', 8))
    y_length_label.grid(row=12, column=0, sticky="e")
    y_length_entry = ttk.Entry(root)
    y_length_entry.grid(row=12, column=1, sticky="we")

    # Create and pack the browse files button
    browse_button = ttk.Button(root, text="Orographic Files", command=browse_files)
    browse_button.grid(row=2, column=0, sticky="we")
    Tooltip(browse_button, "Files must be input one by one, they are automatically copied into the "
                           "folder in which the aermap.inp file is compiled into; if files already "
                           "exist they wont be copied, but will be present in input file")

    compile_button = ttk.Button(root, text="Compile", command=compile_output)
    compile_button.grid(row=13, column=1, sticky="we")
    Tooltip(compile_button, "Choose folder where aermap.inp, receptor.rou, orographic files "
                            "and other associate files will be created")

    # Hide the generate button
    generate_button = ttk.Button(root, text="Generate Output", command=generate_output)
    generate_button.grid_forget()

    # Hide the output text widget
    output_text_label = ttk.Label(root, text="Output Text:")
    output_text_label.grid_forget()
    output_text = tk.Text(root, height=20, width=80)
    output_text.grid_forget()

    root.mainloop()


def app2():
    pointsource_entries = []
    polygon_area_source_entries = []
    manual_polygon_area_source_entries = []

    def generate_output():
        output_text_content = ""
        # CO section
        output_text_content += "CO STARTING\n"
        if title_entry.get():
            output_text_content += "CO TITLEONE " + title_entry.get() + "\n"
        output_text_content += "CO MODELOPT DFAULT CONC\n"
        if time1_entry.get() and time2_entry.get() and time3_entry.get():
            output_text_content += (
                    "CO AVERTIME " + time1_entry.get() + " " +
                    time2_entry.get() + " " + time3_entry.get() + "\n"
            )
        elif time1_entry.get() and time2_entry.get():
            output_text_content += "CO AVERTIME " + time1_entry.get() + " " + time2_entry.get() + "\n"
        elif time1_entry.get():
            output_text_content += "CO AVERTIME " + time1_entry.get() + "\n"
        if pollutant_entry.get():
            output_text_content += "CO POLLUTID " + pollutant_entry.get() + "\n"
        if flagpole_entry.get():
            output_text_content += "CO FLAGPOLE " + flagpole_entry.get() + "\n"
        output_text_content += "CO RUNORNOT RUN\n"
        output_text_content += "CO FINISHED\n\n"

        output_text_content += "SO STARTING\n"
        output_text_content += "SO ELEVUNIT METERS\n"

        point_location_content = ""
        polygon_location_content = ""
        manual_polygon_location_content = ""
        point_srcparam_content = ""
        polygon_srcparam_content = ""
        manual_polygon_srcparam_content = ""
        vertices_location_content = ""
        manual_vertices_location_content = ""

        if pointsource_entries:
            for i, entry in enumerate(pointsource_entries, start=1):
                lat, lon, ptype, rate, height, temp, vel, diameter = entry
                if (lat.get() and lon.get()):
                    point_location_content += f"SO LOCATION STACK{i} POINT {lon.get()} {lat.get()}"
                    if ptype.get():
                        point_location_content += f" {ptype.get()}"
                    point_location_content += "\n"

                    if rate.get() and height.get() and temp.get() and vel.get() and diameter.get():
                        point_srcparam_content += (
                            f"SO SRCPARAM STACK{i} {rate.get()} {height.get()} {temp.get()} {vel.get()}"
                            f" {diameter.get()}\n"
                        )

        if polygon_area_source_entries:
            for j, polygon_entry in enumerate(polygon_area_source_entries, start=1):
                vertices = polygon_entry["vertices"]
                if vertices:
                    first_vertex = vertices[0]
                    first_lat_entry, first_lon_entry = first_vertex
                    polygon_location_content += (f"SO LOCATION POLYGON{j} AREAPOLY {first_lon_entry.get()} "
                                                 f"{first_lat_entry.get()}\n")
                    vertices_location_content += f"SO AREAVERT POLYGON{j} "
                    for vertex_entry in vertices:
                        lat_entry, lon_entry = vertex_entry
                        vertices_location_content += f"{lon_entry.get()} {lat_entry.get()} "
                    vertices_location_content += "\n"
                    if polygon_entry['rate_entry'].get() and polygon_entry['rheight_entry'].get() and polygon_entry[
                        'nvert_entry'].get() and polygon_entry['iheight_entry'].get():
                        polygon_srcparam_content += (
                            f"SO SRCPARAM POLYGON{j} {polygon_entry['rate_entry'].get()} "
                            f"{polygon_entry['rheight_entry'].get()} {polygon_entry['nvert_entry'].get()} "
                            f"{polygon_entry['iheight_entry'].get()}\n"
                        )

        if manual_polygon_area_source_entries:
            for k, manual_polygon_entry in enumerate(manual_polygon_area_source_entries, start=1):
                manual_vertices = manual_polygon_entry["manual_vertices"]
                if manual_vertices:
                    manual_first_vertex = manual_vertices[0]
                    manual_first_lat_entry, manual_first_lon_entry = manual_first_vertex
                    manual_polygon_location_content += (f"SO LOCATION MPOLYGON{k} AREAPOLY "
                                                        f"{manual_first_lon_entry.get()}"
                                                        f" {manual_first_lat_entry.get()}\n")
                    manual_vertices_location_content += f"SO AREAVERT MPOLYGON{j} "
                    for manual_vertex_entry in manual_vertices:
                        manual_lat_entry, manual_lon_entry = manual_vertex_entry
                        manual_vertices_location_content += f"{manual_lon_entry.get()} {manual_lat_entry.get()} "
                        manual_vertices_location_content += "\n"
                    if manual_polygon_entry['manual_rate_entry'].get() and manual_polygon_entry[
                        'manual_rheight_entry'].get() and manual_polygon_entry['manual_nvert_entry'].get() and \
                            manual_polygon_entry['manual_iheight_entry'].get():
                        manual_polygon_srcparam_content += (
                            f"SO SRCPARAM MPOLYGON{k} {manual_polygon_entry['manual_rate_entry'].get()} "
                            f"{manual_polygon_entry['manual_rheight_entry'].get()}"
                            f" {manual_polygon_entry['manual_nvert_entry'].get()} "
                            f"{manual_polygon_entry['manual_iheight_entry'].get()}\n"
                        )

        # Concatenate all content
        output_text_content += (
                point_location_content +
                point_srcparam_content +
                polygon_location_content +
                manual_polygon_location_content +
                polygon_srcparam_content +
                manual_polygon_srcparam_content +
                vertices_location_content +
                manual_vertices_location_content
        )

        if group_name_entry.get():
            point_sources = ['STACK' + str(i) for i in range(1, len(pointsource_entries) + 1)]
            polygon_sources = ['POLYGON' + str(j) for j in range(1, len(polygon_area_source_entries) + 1)]
            manual_polygon_sources = ['MPOLYGON' + str(k) for k in
                                      range(1, len(manual_polygon_area_source_entries) + 1)]
            all_sources = point_sources + polygon_sources + manual_polygon_sources
            output_text_content += (
                f"SO SRCGROUP {group_name_entry.get()} {' '.join(all_sources)}\n"
            )

        output_text_content += "SO FINISHED\n\n"

        output_text_content += f"RE STARTING\nRE INCLUDED {chosen_file_entry_map_output.get()}\nRE FINISHED\n\n"

        output_text_content += "ME STARTING\n"
        output_text_content += "ME SURFFILE " + chosen_file_entry_sfc_output.get() + "\n"
        output_text_content += "ME PROFFILE " + chosen_file_entry_prof_output.get() + "\n"
        if station_num_entry.get() and start_year_entry.get():
            output_text_content += f"ME SURFDATA {station_num_entry.get()} {start_year_entry.get()}\n"
        if upper_air_station_num_entry.get() and start_year_upper_air_entry.get():
            output_text_content += (f"ME UAIRDATA {upper_air_station_num_entry.get()} "
                                    f"{start_year_upper_air_entry.get()}\n")
        if base_elevation_entry.get():
            output_text_content += f"ME PROFBASE {base_elevation_entry.get()} METERS\n"
        if start_date_entry.get() and end_date_entry.get():
            output_text_content += f"ME STARTEND {start_date_entry.get()} {end_date_entry.get()}\n"
        output_text_content += "ME FINISHED\n\n"

        output_text_content += "OU STARTING\n"
        if rec_table_entry.get():
            output_text_content += f"OU RECTABLE ALLAVE {rec_table_entry.get()}\n"
        if max_table_entry.get():
            output_text_content += f"OU MAXTABLE ALLAVE {max_table_entry.get()}\n"
        if time1_entry.get() and rank1_hinum_entry.get():
            output_text_content += (
                f"OU RANKFILE {time1_entry.get()} "
                f"{rank1_hinum_entry.get()} RANK{time1_entry.get()}.RNK\n"
            )
        if time2_entry.get() and rank2_hinum_entry.get():
            output_text_content += (
                f"OU RANKFILE {time2_entry.get()} "
                f"{rank2_hinum_entry.get()} RANK{time2_entry.get()}.RNK\n"
            )
        if time3_entry.get() and rank3_hinum_entry.get():
            output_text_content += (
                f"OU RANKFILE {time3_entry.get()} "
                f"{rank3_hinum_entry.get()} RANK{time3_entry.get()}.RNK\n"
            )
        if time1_entry.get() and group_name_entry.get() and max1_value_entry.get():
            output_text_content += (f"OU MAXIFILE {time1_entry.get()} {group_name_entry.get()} "
                                    f"{max1_value_entry.get()} MAX{time1_entry.get()}H_{group_name_entry.get()}.OUT\n"
                                    )
        if time2_entry.get() and group_name_entry.get() and max2_value_entry.get():
            output_text_content += (
                f"OU MAXIFILE {time2_entry.get()} {group_name_entry.get()} "
                f"{max2_value_entry.get()} MAX{time2_entry.get()}H_{group_name_entry.get()}.OUT\n"
            )
        if time3_entry.get() and group_name_entry.get() and max3_value_entry.get():
            output_text_content += (
                f"OU MAXIFILE {time3_entry.get()} {group_name_entry.get()} "
                f"{max3_value_entry.get()} MAX{time3_entry.get()}H_{group_name_entry.get()}.OUT\n"
            )
        if time1_entry.get() and group_name_entry.get():
            output_text_content += (
                f"OU PLOTFILE {time1_entry.get()} {group_name_entry.get()} "
                f"FIRST PLOT{time1_entry.get()}H_{group_name_entry.get()}.PLT\n"
            )
        if time2_entry.get() and group_name_entry.get():
            output_text_content += (
                f"OU PLOTFILE {time2_entry.get()} {group_name_entry.get()} "
                f"FIRST PLOT{time2_entry.get()}H_{group_name_entry.get()}.PLT\n"
            )
        if time3_entry.get() and group_name_entry.get():
            output_text_content += (
                f"OU PLOTFILE {time3_entry.get()} {group_name_entry.get()} "
                f"FIRST PLOT{time3_entry.get()}H_{group_name_entry.get()}.PLT\n"
            )
        output_text_content += "OU FINISHED\n"

        return output_text_content

    def compile_output():
        output_text_content = generate_output()
        folder_path = filedialog.askdirectory()
        if folder_path:
            file_path = os.path.join(folder_path, "aermod.inp")
            with open(file_path, "w") as file:
                file.write(output_text_content)

            # Copy selected files to the destination folder
            for chosen_file_entry in [chosen_file_entry_map_output, chosen_file_entry_sfc_output,
                                      chosen_file_entry_prof_output]:
                filename = chosen_file_entry.get()
                if filename:
                    source_path = os.path.join(folder_path, filename)
                    destination = os.path.join(folder_path, filename)
                    if source_path != destination:
                        shutil.copy(source_path, destination)

            root.destroy()

    def get_clipboard_text():
        try:
            win32clipboard.OpenClipboard()
            clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
            win32clipboard.CloseClipboard()
            return clipboard_data.decode('utf-8')
        except (UnicodeDecodeError, TypeError):
            return ""

    def open_file_dialog_map_output():
        file_path_map_output = filedialog.askopenfilename()
        if file_path_map_output:
            file_name = os.path.basename(file_path_map_output)
            chosen_file_entry_map_output.delete(0, tk.END)
            chosen_file_entry_map_output.insert(0, file_name)

    def open_file_dialog_sfc_output():
        file_path_sfc_output = filedialog.askopenfilename()
        if file_path_sfc_output:
            file_name = os.path.basename(file_path_sfc_output)
            chosen_file_entry_sfc_output.delete(0, tk.END)
            chosen_file_entry_sfc_output.insert(0, file_name)

    def open_file_dialog_prof_output():
        file_path_prof_output = filedialog.askopenfilename()
        if file_path_prof_output:
            file_name = os.path.basename(file_path_prof_output)
            chosen_file_entry_prof_output.delete(0, tk.END)
            chosen_file_entry_prof_output.insert(0, file_name)

    class Tooltip:
        def __init__(self, widget, text):
            self.widget = widget
            self.text = text
            self.tooltip = None
            self.widget.bind("<Enter>", self.enter)
            self.widget.bind("<Leave>", self.leave)

        def enter(self, event=None):
            x, y, _, _ = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 25
            if event:
                x = event.x_root + 10
                y = event.y_root + 10
            self.tooltip = tk.Toplevel(self.widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()

        def leave(self, event=None):
            if self.tooltip:
                self.tooltip.destroy()

    def add_pointsource():
        lat_label = ttk.Label(pointsource_frame, text=f"Northing {len(pointsource_entries) + 1}:", font=('Arial', 8))
        lat_label.grid(row=1, column=len(pointsource_entries) * 4, sticky="e")

        Tooltip(lat_label, "UTM coordinates, up to 4 decimal spots")

        lat_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        lat_entry.grid(row=1, column=len(pointsource_entries) * 4 + 1, sticky="we")

        lon_label = ttk.Label(pointsource_frame, text=f"Easting {len(pointsource_entries) + 1}:", font=('Arial', 8))
        lon_label.grid(row=1, column=len(pointsource_entries) * 4 + 2, sticky="e")

        Tooltip(lon_label, "UTM coordinates, up to 4 decimal spots")

        lon_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        lon_entry.grid(row=1, column=len(pointsource_entries) * 4 + 3, sticky="we")

        ptype_label = ttk.Label(pointsource_frame, text=f"Base Elevation(m) {len(pointsource_entries) + 1}:", font=('Arial', 8))
        ptype_label.grid(row=2, column=len(pointsource_entries) * 4, sticky="e")
        Tooltip(ptype_label, "Elevation of the stack base from sea level")
        ptype_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        ptype_entry.grid(row=2, column=len(pointsource_entries) * 4 + 1, sticky="we")

        rate_label = ttk.Label(pointsource_frame, text=f"Emission Rate(g/s) {len(pointsource_entries) + 1}:",
                               font=('Arial', 8))
        rate_label.grid(row=2, column=len(pointsource_entries) * 4 + 2, sticky="e")
        rate_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        rate_entry.grid(row=2, column=len(pointsource_entries) * 4 + 3, sticky="we")

        height_label = ttk.Label(pointsource_frame, text=f"Stack Height(m) {len(pointsource_entries) + 1}:",
                                 font=('Arial', 8))
        height_label.grid(row=3, column=len(pointsource_entries) * 4, sticky="e")
        height_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        height_entry.grid(row=3, column=len(pointsource_entries) * 4 + 1, sticky="we")

        temp_label = ttk.Label(pointsource_frame, text=f"Exit Temp.(K) {len(pointsource_entries) + 1}:",
                               font=('Arial', 8))
        temp_label.grid(row=3, column=len(pointsource_entries) * 4 + 2, sticky="e")
        temp_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        temp_entry.grid(row=3, column=len(pointsource_entries) * 4 + 3, sticky="we")

        vel_label = ttk.Label(pointsource_frame, text=f"Exit Velocity(m/s) {len(pointsource_entries) + 1}:",
                              font=('Arial', 8))
        vel_label.grid(row=4, column=len(pointsource_entries) * 4, sticky="e")
        vel_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        vel_entry.grid(row=4, column=len(pointsource_entries) * 4 + 1, sticky="we")

        diameter_label = ttk.Label(pointsource_frame, text=f"Stack Diameter(m) {len(pointsource_entries) + 1}:",
                                   font=('Arial', 8))
        diameter_label.grid(row=4, column=len(pointsource_entries) * 4 + 2, sticky="e")
        diameter_entry = ttk.Entry(pointsource_frame, width=9, font=('Arial', 8))
        diameter_entry.grid(row=4, column=len(pointsource_entries) * 4 + 3, sticky="we")

        choose_on_map_button = ttk.Button(pointsource_frame, text="Open map",
                                          command=lambda: open_google_maps_for_pointsource(lat_entry, lon_entry),
                                          style='Custom.TButton')
        pointsource_frame.style = ttk.Style()
        pointsource_frame.style.configure('Custom.TButton',
                                          font=('Arial', 8))

        choose_on_map_button.grid(row=0, column=len(pointsource_entries) * 4)

        Tooltip(choose_on_map_button, "Opens Google Maps; Copied coordinates are automatically converted to UTM "
                                      "and input into the textboxes, Google Earth is opened to display point sources")

        # Add the new entries to the list
        pointsource_entries.append((lat_entry, lon_entry, ptype_entry,
                                    rate_entry, height_entry, temp_entry, vel_entry, diameter_entry))

    def open_google_maps_for_pointsource(lat_entry, lon_entry):
        url = "https://www.google.com/maps"
        webbrowser.open(url)

        def monitor_clipboard():
            last_clipboard_text = get_clipboard_text()
            while True:
                clipboard_text = get_clipboard_text()
                if clipboard_text != last_clipboard_text:
                    last_clipboard_text = clipboard_text
                    try:
                        lat, lon = map(float, clipboard_text.split(','))
                        utm_coords = utm.from_latlon(lat, lon)
                        utm_easting, utm_northing, utm_zone_number, utm_zone_letter = utm_coords
                        if lat_entry.get() == '' and lon_entry.get() == '':
                            lat, lon = map(float, clipboard_text.split(','))
                            update_kmz(lat, lon, 'point')
                            os.startfile("updated_file.kmz")
                            lat_entry.delete(0, 'end')
                            lon_entry.delete(0, 'end')
                            utm_northing_rounded = round(utm_northing, 4)
                            utm_easting_rounded = round(utm_easting, 4)
                            lat_entry.insert(0, utm_northing_rounded)
                            lon_entry.insert(0, utm_easting_rounded)
                    except ValueError:
                        print("Invalid coordinates format in clipboard")
                time.sleep(1)

        clipboard_thread = threading.Thread(target=monitor_clipboard)
        clipboard_thread.daemon = True
        clipboard_thread.start()

    def add_polygon_area_source():
        polygon_entry = {"vertices": [], "rate_entry": "", "rheight_entry": "", "nvert_entry": "", "iheight_entry": ""}
        polygon_area_source_entries.append(polygon_entry)

        add_vertex_button = ttk.Button(polygon_area_source_frame,
                                       text="Add Vertex",
                                       command=lambda: add_polygon_area_vertex(polygon_entry, lat, lon),
                                       style='Custom.TButton')
        polygon_area_source_frame.style = ttk.Style()
        polygon_area_source_frame.style.configure('Custom.TButton',
                                                  font=('Arial', 8))

        rate_label = ttk.Label(polygon_area_source_frame,
                               text=f"Emission Rate(g/s) {len(polygon_area_source_entries)}:",
                               font=('Arial', 8))
        rate_label.grid(row=3 * len(polygon_area_source_entries) + 1, column=0, sticky="w")
        rate_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        rate_entry.grid(row=3 * len(polygon_area_source_entries) + 1, column=1, sticky="w")
        polygon_entry["rate_entry"] = rate_entry

        rheight_label = ttk.Label(polygon_area_source_frame,
                                  text=f"Release Height(m) {len(polygon_area_source_entries)}:",
                                  font=('Arial', 8))
        rheight_label.grid(row=3 * len(polygon_area_source_entries) + 1, column=2, sticky="w")
        rheight_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        rheight_entry.grid(row=3 * len(polygon_area_source_entries) + 1, column=3, sticky="w")
        polygon_entry["rheight_entry"] = rheight_entry

        nvert_label = ttk.Label(polygon_area_source_frame, text=f"N° Vertices {len(polygon_area_source_entries)}:",
                                font=('Arial', 8))
        nvert_label.grid(row=3 * len(polygon_area_source_entries) + 2, column=0, sticky="w")
        nvert_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        nvert_entry.grid(row=3 * len(polygon_area_source_entries) + 2, column=1, sticky="w")
        polygon_entry["nvert_entry"] = nvert_entry

        iheight_label = ttk.Label(polygon_area_source_frame,
                                  text=f"Initial Source Height(m) {len(polygon_area_source_entries)}:",
                                  font=('Arial', 8))
        iheight_label.grid(row=3 * len(polygon_area_source_entries) + 2, column=2, sticky="w")
        iheight_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        iheight_entry.grid(row=3 * len(polygon_area_source_entries) + 2, column=3, sticky="w")
        polygon_entry["iheight_entry"] = iheight_entry

        Tooltip(iheight_label, "Optional")

        choose_on_map_button = ttk.Button(polygon_area_source_frame, text="Open map",
                                          command=lambda: open_google_maps_for_polygon(polygon_entry),
                                          style='Custom.TButton')
        polygon_area_source_frame.style = ttk.Style()
        polygon_area_source_frame.style.configure('Custom.TButton',
                                                  font=('Arial', 8))

        choose_on_map_button.grid(row=3 * len(polygon_area_source_entries), column=1)

        Tooltip(choose_on_map_button, "Opens Google Maps; Copied coordinates are"
                                      "automatically converted to UTM,create a new vertex and insert the values. "
                                      "Google Earth is opened to display the polygons")

    def add_polygon_area_vertex(polygon_entry, lat, lon):
        global new_vertex_row

        new_vertex_row = 1 + len(polygon_entry['vertices'])

        if len(polygon_entry['vertices']) == 0:
            vertex_label = ttk.Label(polygon_area_source_frame,
                                     text=f"Area {len(polygon_area_source_entries)} Vertices:",
                                     font=('Arial', 8))
            vertex_label.grid(row=0, column=4 * len(polygon_area_source_entries) + 1, columnspan=2, sticky="w")

        lat_label = ttk.Label(polygon_area_source_frame, text=f"Northing {len(polygon_entry['vertices']) + 1}:",
                              font=('Arial', 8))
        lat_label.grid(row=new_vertex_row, column=4 * len(polygon_area_source_entries) + 1, sticky="w")

        Tooltip(lat_label, "Up to 4 decimal spots")

        lat_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        lat_entry.grid(row=new_vertex_row, column=4 * len(polygon_area_source_entries) + 2, sticky="w")
        lat_entry.insert(0, lat)

        lon_label = ttk.Label(polygon_area_source_frame, text=f"Easting {len(polygon_entry['vertices']) + 1}:",
                              font=('Arial', 8))
        lon_label.grid(row=new_vertex_row, column=4 * len(polygon_area_source_entries) + 3, sticky="w")

        Tooltip(lon_label, "Up to 4 decimal spots")

        lon_entry = ttk.Entry(polygon_area_source_frame, width=9, font=('Arial', 8))
        lon_entry.grid(row=new_vertex_row, column=4 * len(polygon_area_source_entries) + 4, sticky="w")
        lon_entry.insert(0, lon)

        polygon_entry['vertices'].append((lat_entry, lon_entry))

    def open_google_maps_for_polygon(polygon_entry):
        global current_geometry_type
        current_geometry_type = 'polygon'
        url = "https://www.google.com/maps"
        webbrowser.open(url)

        initial_source_count = len(polygon_area_source_entries)

        def monitor_clipboard(polygon_entry, initial_source_count):
            last_clipboard_text = get_clipboard_text()
            while True:
                clipboard_text = get_clipboard_text()
                if clipboard_text != last_clipboard_text:
                    last_clipboard_text = clipboard_text
                    if len(polygon_area_source_entries) > initial_source_count:
                        return
                    try:
                        lat, lon = map(float, clipboard_text.split(','))
                        update_kmz(lat, lon, 'polygon')
                        os.startfile("updated_file.kmz")
                        utm_coords = utm.from_latlon(lat, lon)
                        utm_easting, utm_northing, utm_zone_number, utm_zone_letter = utm_coords
                        lat = round(utm_northing, 4)
                        lon = round(utm_easting, 4)
                        add_polygon_area_vertex(polygon_entry, lat, lon)
                    except ValueError:
                        print("Invalid coordinates format in clipboard")
                time.sleep(1)

        clipboard_thread = threading.Thread(target=monitor_clipboard, args=(polygon_entry, initial_source_count))
        clipboard_thread.daemon = True
        clipboard_thread.start()

    kml = simplekml.Kml()
    polygon_vertices = []

    def update_kmz(lat, lon, geometry_type):
        if geometry_type == 'point':
            kml.newpoint(name="Point Source", coords=[(lon, lat)])
        elif geometry_type == 'polygon':
            polygon_vertices.append((lon, lat))
            if len(polygon_vertices) >= 4:
                polygon = kml.newpolygon(name="Polygon Area", outerboundaryis=polygon_vertices)
                polygon.style.linestyle.color = 'ff0000ff'
        kml.save("updated_file.kmz")

    current_geometry_type = None

    def delete_polygon_vertices():
        polygon_vertices.clear()

    def manual_add_polygon_area_source():
        manual_polygon_entry = {"manual_vertices": [], "manual_rate_entry": "", "manual_rheight_entry": "",
                                "manual_nvert_entry": "", "manual_iheight_entry": ""}
        manual_polygon_area_source_entries.append(manual_polygon_entry)

        manual_add_vertex_button = ttk.Button(manual_polygon_area_source_frame,
                                              text="Add Vertex",
                                              command=lambda: manual_add_polygon_area_vertex(manual_polygon_entry),
                                              style='Custom.TButton')
        manual_polygon_area_source_frame.style = ttk.Style()
        manual_polygon_area_source_frame.style.configure('Custom.TButton',
                                                         font=('Arial', 8))

        manual_add_vertex_button.grid(row=3 * len(manual_polygon_area_source_entries), column=0, sticky="w")

        manual_rate_label = ttk.Label(manual_polygon_area_source_frame,
                                      text=f"Emission Rate(g/s) {len(manual_polygon_area_source_entries)}:",
                                      font=('Arial', 8))

        manual_rate_label.grid(row=3 * len(manual_polygon_area_source_entries) + 1, column=0, sticky="w")
        manual_rate_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_rate_entry.grid(row=3 * len(manual_polygon_area_source_entries) + 1, column=1, sticky="w")
        manual_polygon_entry["manual_rate_entry"] = manual_rate_entry

        manual_rheight_label = ttk.Label(manual_polygon_area_source_frame,
                                         text=f"Release Height(m) {len(manual_polygon_area_source_entries)}:",
                                         font=('Arial', 8))

        manual_rheight_label.grid(row=3 * len(manual_polygon_area_source_entries) + 1, column=2, sticky="w")
        manual_rheight_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_rheight_entry.grid(row=3 * len(manual_polygon_area_source_entries) + 1, column=3, sticky="w")
        manual_polygon_entry["manual_rheight_entry"] = manual_rheight_entry

        manual_nvert_label = ttk.Label(manual_polygon_area_source_frame,
                                       text=f"N° Vertices {len(manual_polygon_area_source_entries)}:",
                                       font=('Arial', 8))

        manual_nvert_label.grid(row=3 * len(manual_polygon_area_source_entries) + 2, column=0, sticky="w")
        manual_nvert_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_nvert_entry.grid(row=3 * len(manual_polygon_area_source_entries) + 2, column=1, sticky="w")
        manual_polygon_entry["manual_nvert_entry"] = manual_nvert_entry

        manual_iheight_label = ttk.Label(manual_polygon_area_source_frame,
                                         text=f"Initial Source Height(m) {len(manual_polygon_area_source_entries)}:",
                                         font=('Arial', 8))
        manual_iheight_label.grid(row=3 * len(manual_polygon_area_source_entries) + 2, column=2, sticky="w")
        manual_iheight_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_iheight_entry.grid(row=3 * len(manual_polygon_area_source_entries) + 2, column=3, sticky="w")
        manual_polygon_entry["manual_iheight_entry"] = manual_iheight_entry

        Tooltip(manual_iheight_label, "Optional")

    def manual_add_polygon_area_vertex(manual_polygon_entry):
        global manual_new_vertex_row

        manual_new_vertex_row = 1 + len(manual_polygon_entry['manual_vertices'])

        if len(manual_polygon_entry['manual_vertices']) == 0:
            manual_vertex_label = ttk.Label(manual_polygon_area_source_frame,
                                            text=f"Area {len(manual_polygon_area_source_entries)} Vertices:",
                                            font=('Arial', 8))
            manual_vertex_label.grid(row=0, column=4 * len(manual_polygon_area_source_entries) + 1,
                                     columnspan=2, sticky="w")

        manual_lat_label = ttk.Label(manual_polygon_area_source_frame,
                                     text=f"Northing {len(manual_polygon_entry['manual_vertices']) + 1}:",
                                     font=('Arial', 8))
        manual_lat_label.grid(row=manual_new_vertex_row, column=4 * len(manual_polygon_area_source_entries) + 1,
                              sticky="w")

        Tooltip(manual_lat_label, "Up to 4 decimal spots, UTM coordinates")

        manual_lat_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_lat_entry.grid(row=manual_new_vertex_row, column=4 * len(manual_polygon_area_source_entries) + 2,
                              sticky="w")

        manual_lon_label = ttk.Label(manual_polygon_area_source_frame,
                                     text=f"Easting {len(manual_polygon_entry['manual_vertices']) + 1}:",
                                     font=('Arial', 8))
        manual_lon_label.grid(row=manual_new_vertex_row, column=4 * len(manual_polygon_area_source_entries) + 3,
                              sticky="w")

        Tooltip(manual_lon_label, "Up to 4 decimal spots, UTM coordinates")

        manual_lon_entry = ttk.Entry(manual_polygon_area_source_frame, width=9, font=('Arial', 8))
        manual_lon_entry.grid(row=manual_new_vertex_row, column=4 * len(manual_polygon_area_source_entries) + 4,
                              sticky="w")

        manual_polygon_entry['manual_vertices'].append((manual_lat_entry, manual_lon_entry))

    root = tk.Tk()
    root.title("AERMOD Input File Generator")

    root.geometry("1000x1000")

    frame = ttk.Frame(root)
    frame.grid(row=0, column=0, sticky="nsew")

    canvas = tk.Canvas(frame)
    scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    content_frame = ttk.Frame(canvas)

    content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    canvas.create_window((0, 0), window=content_frame, anchor="nw")
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    def _on_mousewheel(event):
        if event.state & 0x4:  # Check if Ctrl key is pressed
            canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    title_label = ttk.Label(content_frame, text="Title:", font=('Arial', 8))
    title_label.grid(row=0, column=0, sticky="e")
    title_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    title_entry.grid(row=0, column=1, sticky="w")

    time1_label = ttk.Label(content_frame, text="Time period 1 (h):", font=('Arial', 8))
    time1_label.grid(row=1, column=0, sticky="e")

    Tooltip(time1_label, "Defines the first averaging period")

    time1_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    time1_entry.grid(row=1, column=1, sticky="w")

    time2_label = ttk.Label(content_frame, text="Time period 2 (h):", font=('Arial', 8))
    time2_label.grid(row=2, column=0, sticky="e")

    Tooltip(time2_label, "Defines the second time point when the analysis is done")

    time2_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    time2_entry.grid(row=2, column=1, sticky="w")

    time3_label = ttk.Label(content_frame, text="Time period 3 (h):", font=('Arial', 8))
    time3_label.grid(row=3, column=0, sticky="e")

    Tooltip(time3_label, "Defines the third time point when the analysis is done")

    time3_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    time3_entry.grid(row=3, column=1, sticky="w")

    pollutant_label = ttk.Label(content_frame, text="Pollutant:", font=('Arial', 8))
    pollutant_label.grid(row=4, column=0, sticky="e")

    Tooltip(pollutant_label, "SO2 CO NOX NO2 TSP PM10 PM2.5 LEAD OTHER")

    pollutant_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    pollutant_entry.grid(row=4, column=1, sticky="w")

    flagpole_label = ttk.Label(content_frame, text="Flagpole/Receptor Height (m):", font=('Arial', 8))
    flagpole_label.grid(row=5, column=0, sticky="e")
    flagpole_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    flagpole_entry.grid(row=5, column=1, sticky="w")

    pointsource_frame = ttk.Frame(content_frame)
    pointsource_frame.grid(row=6, column=0, columnspan=2)

    add_pointsource_button = ttk.Button(content_frame, text="Add Point Source", command=add_pointsource,
                                        style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))

    add_pointsource_button.grid(row=7, column=0, columnspan=2)

    polygon_area_source_frame = ttk.Frame(content_frame)
    polygon_area_source_frame.grid(row=8, column=0, columnspan=2)

    add_polygon_area_source_button = ttk.Button(content_frame,
                                                text="Add Polygon Area Source With Google Maps",
                                                command=lambda: (add_polygon_area_source(), delete_polygon_vertices()),
                                                style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))

    add_polygon_area_source_button.grid(row=9, column=0, columnspan=2)

    manual_polygon_area_source_frame = ttk.Frame(content_frame)
    manual_polygon_area_source_frame.grid(row=10, column=0, columnspan=2)

    manual_add_polygon_area_source_button = ttk.Button(content_frame, text="Add Polygon Area Source Manually",
                                                       command=manual_add_polygon_area_source, style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))

    manual_add_polygon_area_source_button.grid(row=11, column=0, columnspan=2)

    group_name_label = ttk.Label(content_frame, text="Group name:", font=('Arial', 8))
    group_name_label.grid(row=12, column=0, sticky="e")

    Tooltip(group_name_label, "Up to 8 alphanumeric characters, is used along with Time periods to create "
                              "plotfile name")

    group_name_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    group_name_entry.grid(row=12, column=1, sticky="w")

    open_file_button_map_output = ttk.Button(content_frame, text="Choose Receptor File",
                                             command=open_file_dialog_map_output, style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))
    open_file_button_map_output.grid(row=13, column=0, sticky="e")

    Tooltip(open_file_button_map_output, ".rec/.rou file is located in same location as aermap input file")

    chosen_file_entry_map_output = ttk.Entry(content_frame, width=50, font=('Arial', 8))
    chosen_file_entry_map_output.grid(row=13, column=1, sticky="w")

    open_file_button_sfc_output = ttk.Button(content_frame, text="Choose Surface Meteo Data File",
                                             command=open_file_dialog_sfc_output, style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))
    open_file_button_sfc_output.grid(row=14, column=0, sticky="e")

    Tooltip(open_file_button_sfc_output, ".sfc file is located in same location as aermet input file")

    chosen_file_entry_sfc_output = ttk.Entry(content_frame, width=50, font=('Arial', 8))
    chosen_file_entry_sfc_output.grid(row=14, column=1, sticky="w")

    open_file_button_prof_output = ttk.Button(content_frame, text="Choose Meteo Profile Data File",
                                              command=open_file_dialog_prof_output, style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))
    open_file_button_prof_output.grid(row=15, column=0, sticky="e")

    Tooltip(open_file_button_prof_output, ".pfl file is located in same location as aermet input file")

    chosen_file_entry_prof_output = ttk.Entry(content_frame, width=50, font=('Arial', 8))
    chosen_file_entry_prof_output.grid(row=15, column=1, sticky="w")

    station_num_label = ttk.Label(content_frame, text="Surface Data Station Number:", font=('Arial', 8))
    station_num_label.grid(row=16, column=0, sticky="e")
    station_num_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    station_num_entry.grid(row=16, column=1, sticky="w")

    start_year_label = ttk.Label(content_frame, text="Start Year:", font=('Arial', 8))
    start_year_label.grid(row=17, column=0, sticky="e")

    Tooltip(start_year_label, "Starting year of the data, regardless of when the analysis starts")

    start_year_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    start_year_entry.grid(row=17, column=1, sticky="w")

    upper_air_station_num_label = ttk.Label(content_frame, text="Upper Air Data Station Number:", font=('Arial', 8))
    upper_air_station_num_label.grid(row=18, column=0, sticky="e")
    upper_air_station_num_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    upper_air_station_num_entry.grid(row=18, column=1, sticky="w")

    start_year_upper_air_label = ttk.Label(content_frame, text="Start Year (Upper Air):", font=('Arial', 8))
    start_year_upper_air_label.grid(row=19, column=0, sticky="e")

    Tooltip(start_year_upper_air_label, "Starting year of the data, regardless of when the analysis starts")

    start_year_upper_air_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    start_year_upper_air_entry.grid(row=19, column=1, sticky="w")

    base_elevation_label = ttk.Label(content_frame, text="Base Elevation (m):", font=('Arial', 8))
    base_elevation_label.grid(row=20, column=0, sticky="e")
    base_elevation_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    base_elevation_entry.grid(row=20, column=1, sticky="w")

    start_date_label = ttk.Label(content_frame, text="Start Date (YYYY MM DD):", font=('Arial', 8))
    start_date_label.grid(row=21, column=0, sticky="e")

    Tooltip(start_date_label, "Starting date of the analysis (YYYY MM DD)")

    start_date_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    start_date_entry.grid(row=21, column=1, sticky="w")

    end_date_label = ttk.Label(content_frame, text="End Date (YYYY MM DD):", font=('Arial', 8))
    end_date_label.grid(row=22, column=0, sticky="e")

    Tooltip(end_date_label, "End date of the analysis (YYYY MM DD)")

    end_date_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    end_date_entry.grid(row=22, column=1, sticky="w")

    rec_table_label = ttk.Label(content_frame, text="Rec Table (1ST 2ND 3RD):", font=('Arial', 8))
    rec_table_label.grid(row=23, column=0, sticky="e")

    Tooltip(rec_table_label, "Select highest values for output tables")

    rec_table_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    rec_table_entry.grid(row=23, column=1, sticky="w")

    max_table_label = ttk.Label(content_frame, text="Max Table N° of entries:", font=('Arial', 8))
    max_table_label.grid(row=24, column=0, sticky="e")

    Tooltip(max_table_label, "Input N° of values for the maximum values of all averaging periods")

    max_table_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    max_table_entry.grid(row=24, column=1, sticky="w")

    rank1_label = ttk.Label(content_frame, text="Hinum for 1st Time period Rank:", font=('Arial', 8))
    rank1_label.grid(row=25, column=0, sticky="e")

    Tooltip(rank1_label, "Input N° of max values included in output table of the first averaging period")

    rank1_hinum_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    rank1_hinum_entry.grid(row=25, column=1, sticky="w")

    rank2_label = ttk.Label(content_frame, text="Hinum for 2nd Time period Rank:", font=('Arial', 8))
    rank2_label.grid(row=26, column=0, sticky="e")

    Tooltip(rank2_label, "Input N° of max values included in output table of the second averaging period")

    rank2_hinum_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    rank2_hinum_entry.grid(row=26, column=1, sticky="w")

    rank3_label = ttk.Label(content_frame, text="Hinum for 3rd Time period Rank:", font=('Arial', 8))
    rank3_label.grid(row=27, column=0, sticky="e")

    Tooltip(rank3_label, "Input N° of max values included in output table of the third averaging period")

    rank3_hinum_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    rank3_hinum_entry.grid(row=27, column=1, sticky="w")

    max1_label = ttk.Label(content_frame, text="Threshold 1st Time period:", font=('Arial', 8))
    max1_label.grid(row=28, column=0, sticky="e")
    max1_value_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    max1_value_entry.grid(row=28, column=1, sticky="w")

    max2_label = ttk.Label(content_frame, text="Threshold 2nd Time period:", font=('Arial', 8))
    max2_label.grid(row=29, column=0, sticky="e")
    max2_value_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    max2_value_entry.grid(row=29, column=1, sticky="w")

    max3_label = ttk.Label(content_frame, text="Threshold 3rd Time period:", font=('Arial', 8))
    max3_label.grid(row=30, column=0, sticky="e")
    max3_value_entry = ttk.Entry(content_frame, width=9, font=('Arial', 8))
    max3_value_entry.grid(row=30, column=1, sticky="w")

    compile_button = ttk.Button(content_frame, text="Compile", command=compile_output,
                                style='Custom.TButton')
    root.style = ttk.Style()
    root.style.configure('Custom.TButton', font=('Arial', 8))

    compile_button.grid(row=31, column=0, columnspan=2)
    Tooltip(compile_button, "Name the file aermod.inp")

    root.mainloop()


def app4():
    def browse_files():
        filename = filedialog.askopenfilename()
        if filename:
            new_entry = ttk.Entry(datafile_frame)
            new_entry.grid(row=len(datafile_entries), column=1, sticky="we")
            new_entry.insert(0, filename)
            datafile_entries.append(new_entry)

    def generate_output1():
        output_text_content1 = ""
        output_text_content1 += "version=" + version_entry.get() + "\n"
        output_text_content1 += "origin=UTM\n"
        output_text_content1 += "easting=0\n"
        output_text_content1 += "northing=0\n"
        output_text_content1 += "utmZone=" + utm_entry.get() + "\n"
        output_text_content1 += hemisphere_combobox.get() + "=true\n"
        output_text_content1 += "originLatitude =0\n"
        output_text_content1 += "originLongitude =0\n"
        output_text_content1 += "altitudeChoice = relativeToGround\n"
        output_text_content1 += "altitude=0\n"
        output_text_content1 += (
            f"PlotFileName =PLOT{time1_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content1 += "SourceDisplayInputFileName=aermod.inp\n"
        output_text_content1 += (
            f"OutputFileNameBase =PLOT{time1_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content1 += (
            f"NameDisplayedInGoogleEarth=PLOT{time1_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content1 += "sDisableProgressMeter              = false\n"
        output_text_content1 += "sDisableEarthBrowser               = true\n"
        output_text_content1 += "IconScale     = 0.40\n"
        output_text_content1 += "sIconSetChoice=redBlue\n"
        output_text_content1 += "minbin=" + min_bin_entry.get() + "\n"
        output_text_content1 += "maxbin=" + max_bin_entry.get() + "\n"
        output_text_content1 += "binningChoice =" + binningchoice_combobox.get() + "\n"
        output_text_content1 += "customBinningElevenLevels=na\n"
        output_text_content1 += (
            "contourLegendTitleHTML =C&nbsp;O&nbsp;N&nbsp;C&nbsp;E&nbsp;N&nbsp;T&nbsp;R&nbsp;A&nbsp;"
            "T&nbsp;I&nbsp;O&nbsp;N&nbsp;S\n")
        output_text_content1 += "numberOfGridCols                   =" + gridcols_entry.get() + "\n"
        output_text_content1 += "numberOfGridRows                   =" + gridrows_entry.get() + "\n"
        output_text_content1 += "numberOfTimesToSmoothContourSurface =" + smooth_entry.get() + "\n"
        output_text_content1 += "makeContours                        =" + contour_combobox.get() + "\n"
        output_text_content1 += "contourExtension =  9999999\n"
        output_text_content1 += "makeGradients                       =" + gradient_combobox.get() + "\n"
        output_text_content1 += "gradientExtension=  9999999\n"
        output_text_content1 += "gradientMaxBin=" + max_bin_entry.get() + "\n"
        output_text_content1 += "gradientMinBin=" + min_bin_entry.get() + "\n"
        output_text_content1 += "gradientBinningChoice=" + gradientbinningchoice_combobox.get() + "\n"
        output_text_content1 += "customGradBinElevenLevels=na\n"
        output_text_content1 += "gradientLegendTitleHTML=Gradient&nbsp;Magnitudes\n"
        output_text_content1 += "provideEvenlySpacedInterpolatedGrid = false\n"

        return output_text_content1

    def generate_output2():
        output_text_content2 = ""
        output_text_content2 += "version=" + version_entry.get() + "\n"
        output_text_content2 += "origin=UTM\n"
        output_text_content2 += "easting=0\n"
        output_text_content2 += "northing=0\n"
        output_text_content2 += "utmZone=" + utm_entry.get() + "\n"
        output_text_content2 += hemisphere_combobox.get() + "=true\n"
        output_text_content2 += "originLatitude =0\n"
        output_text_content2 += "originLongitude =0\n"
        output_text_content2 += "altitudeChoice = relativeToGround\n"
        output_text_content2 += "altitude=0\n"
        output_text_content2 += (
            f"PlotFileName =PLOT{time2_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content2 += "SourceDisplayInputFileName=aermod.inp\n"
        output_text_content2 += (
            f"OutputFileNameBase =PLOT{time2_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content2 += (
            f"NameDisplayedInGoogleEarth=PLOT{time2_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content2 += "sDisableProgressMeter              = false\n"
        output_text_content2 += "sDisableEarthBrowser               = true\n"
        output_text_content2 += "IconScale     = 0.40\n"
        output_text_content2 += "sIconSetChoice=redBlue\n"
        output_text_content2 += "minbin=" + min_bin_entry.get() + "\n"
        output_text_content2 += "maxbin=" + max_bin_entry.get() + "\n"
        output_text_content2 += "binningChoice =" + binningchoice_combobox.get() + "\n"
        output_text_content2 += "customBinningElevenLevels=na\n"
        output_text_content2 += (
            "contourLegendTitleHTML =C&nbsp;O&nbsp;N&nbsp;C&nbsp;E&nbsp;N&nbsp;T&nbsp;R&nbsp;A&nbsp;"
            "T&nbsp;I&nbsp;O&nbsp;N&nbsp;S\n")
        output_text_content2 += "numberOfGridCols                   =" + gridcols_entry.get() + "\n"
        output_text_content2 += "numberOfGridRows                   =" + gridrows_entry.get() + "\n"
        output_text_content2 += "numberOfTimesToSmoothContourSurface =" + smooth_entry.get() + "\n"
        output_text_content2 += "makeContours                        =" + contour_combobox.get() + "\n"
        output_text_content2 += "contourExtension =  9999999\n"
        output_text_content2 += "makeGradients                       =" + gradient_combobox.get() + "\n"
        output_text_content2 += "gradientExtension=  9999999\n"
        output_text_content2 += "gradientMaxBin=" + max_bin_entry.get() + "\n"
        output_text_content2 += "gradientMinBin=" + min_bin_entry.get() + "\n"
        output_text_content2 += "gradientBinningChoice=" + gradientbinningchoice_combobox.get() + "\n"
        output_text_content2 += "customGradBinElevenLevels=na\n"
        output_text_content2 += "gradientLegendTitleHTML=Gradient&nbsp;Magnitudes\n"
        output_text_content2 += "provideEvenlySpacedInterpolatedGrid = false\n"

        return output_text_content2

    def generate_output3():
        output_text_content3 = ""
        output_text_content3 += "version=" + version_entry.get() + "\n"
        output_text_content3 += "origin=UTM\n"
        output_text_content3 += "easting=0\n"
        output_text_content3 += "northing=0\n"
        output_text_content3 += "utmZone=" + utm_entry.get() + "\n"
        output_text_content3 += hemisphere_combobox.get() + "=true\n"
        output_text_content3 += "originLatitude =0\n"
        output_text_content3 += "originLongitude =0\n"
        output_text_content3 += "altitudeChoice = relativeToGround\n"
        output_text_content3 += "altitude=0\n"
        output_text_content3 += (
            f"PlotFileName =PLOT{time3_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content3 += "SourceDisplayInputFileName=aermod.inp\n"
        output_text_content3 += (
            f"OutputFileNameBase =PLOT{time3_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content3 += (
            f"NameDisplayedInGoogleEarth=PLOT{time3_entry.get()}H_{group_name_entry.get()}.PLT\n")
        output_text_content3 += "sDisableProgressMeter              = false\n"
        output_text_content3 += "sDisableEarthBrowser               = true\n"
        output_text_content3 += "IconScale     = 0.40\n"
        output_text_content3 += "sIconSetChoice=redBlue\n"
        output_text_content3 += "minbin=" + min_bin_entry.get() + "\n"
        output_text_content3 += "maxbin=" + max_bin_entry.get() + "\n"
        output_text_content3 += "binningChoice =" + binningchoice_combobox.get() + "\n"
        output_text_content3 += "customBinningElevenLevels=na\n"
        output_text_content3 += (
            "contourLegendTitleHTML =C&nbsp;O&nbsp;N&nbsp;C&nbsp;E&nbsp;N&nbsp;T&nbsp;R&nbsp;A&nbsp;"
            "T&nbsp;I&nbsp;O&nbsp;N&nbsp;S\n")
        output_text_content3 += "numberOfGridCols                   =" + gridcols_entry.get() + "\n"
        output_text_content3 += "numberOfGridRows                   =" + gridrows_entry.get() + "\n"
        output_text_content3 += "numberOfTimesToSmoothContourSurface =" + smooth_entry.get() + "\n"
        output_text_content3 += "makeContours                        =" + contour_combobox.get() + "\n"
        output_text_content3 += "contourExtension =  9999999\n"
        output_text_content3 += "makeGradients                       =" + gradient_combobox.get() + "\n"
        output_text_content3 += "gradientExtension=  9999999\n"
        output_text_content3 += "gradientMaxBin=" + max_bin_entry.get() + "\n"
        output_text_content3 += "gradientMinBin=" + min_bin_entry.get() + "\n"
        output_text_content3 += "gradientBinningChoice=" + gradientbinningchoice_combobox.get() + "\n"
        output_text_content3 += "customGradBinElevenLevels=na\n"
        output_text_content3 += "gradientLegendTitleHTML=Gradient&nbsp;Magnitudes\n"
        output_text_content3 += "provideEvenlySpacedInterpolatedGrid = false\n"

        return output_text_content3

    def compile_output():
        folder_path = filedialog.askdirectory()

        for i in range(1, 4):
            subfolder_path = os.path.join(folder_path, f"aerplot{i}")
            os.makedirs(subfolder_path, exist_ok=True)

            shutil.copy(os.path.join(folder_path, "aermod.inp"), subfolder_path)
            shutil.copy(os.path.join(folder_path, "aermod.out"), subfolder_path)
            shutil.copy(config.get("aerplot_path"), subfolder_path)

        output_text_content1 = generate_output1()
        output_text_content2 = generate_output2()
        output_text_content3 = generate_output3()

        for i, output_text_content in enumerate([output_text_content1, output_text_content2, output_text_content3],
                                                start=1):
            with open(os.path.join(folder_path, f"aerplot{i}", "aerplot.inp"), "w") as file:
                file.write(output_text_content)

        plot_files = [f"PLOT{time_entry.get()}H_{group_name_entry.get()}.PLT" for time_entry in
                      [time1_entry, time2_entry, time3_entry]]
        for i, plot_file in enumerate(plot_files, start=1):
            src_plot_path = os.path.join(folder_path, plot_file)
            dest_plot_path = os.path.join(folder_path, f"aerplot{i}", plot_file)
            shutil.copy(src_plot_path, dest_plot_path)

        root.destroy()

    class Tooltip:
        def __init__(self, widget, text):
            self.widget = widget
            self.text = text
            self.tooltip = None
            self.widget.bind("<Enter>", self.enter)
            self.widget.bind("<Leave>", self.leave)

        def enter(self, event=None):
            x, y, _, _ = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y += self.widget.winfo_rooty() + 25
            if event:
                x = event.x_root + 10
                y = event.y_root + 10
            self.tooltip = tk.Toplevel(self.widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()

        def leave(self, event=None):
            if self.tooltip:
                self.tooltip.destroy()

    def get_clipboard_text():
        try:
            win32clipboard.OpenClipboard()
            clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
            win32clipboard.CloseClipboard()
            return clipboard_data.decode('utf-8')
        except (UnicodeDecodeError, TypeError):
            return ""

    def open_google_maps_for_origin(originlat_entry, originlon_entry):
        url = "https://www.google.com/maps"
        webbrowser.open(url)

        def monitor_clipboard():
            last_clipboard_text = get_clipboard_text()
            while True:
                clipboard_text = get_clipboard_text()
                if clipboard_text != last_clipboard_text:
                    last_clipboard_text = clipboard_text
                    try:
                        lat, lon = map(float, clipboard_text.split(','))
                        utm_coords = utm.from_latlon(lat, lon)
                        utm_easting, utm_northing, utm_zone_number, utm_zone_letter = utm_coords
                        if originlat_entry.get() == '' and originlon_entry.get() == '':
                            originlat_entry.delete(0, 'end')
                            originlon_entry.delete(0, 'end')
                            utm_entry.delete(0, 'end')
                            utm_northing_rounded = round(utm_northing, 4)
                            utm_easting_rounded = round(utm_easting, 4)
                            originlat_entry.insert(0, utm_northing_rounded)
                            originlon_entry.insert(0, utm_easting_rounded)
                            utm_entry.insert(0, utm_zone_number)
                    except ValueError:
                        print("Invalid coordinates format in clipboard")
                time.sleep(1)

        clipboard_thread = threading.Thread(target=monitor_clipboard)
        clipboard_thread.daemon = True
        clipboard_thread.start()

    root = tk.Tk()
    root.title("AERPLOT Input File Generator")

    version_label = ttk.Label(root, text="Version:", font=('Arial', 8))
    version_label.grid(row=0, column=0, sticky="e")
    Tooltip(version_label, "Default=2")
    version_entry = ttk.Entry(root, font=('Arial', 8))
    version_entry.grid(row=0, column=1, sticky="we")

    choose_on_map_button = ttk.Button(root, text="Open map for UTM zone",
                                      command=lambda: open_google_maps_for_origin(originlat_entry, originlon_entry))
    choose_on_map_button.grid(row=2, column=1, sticky="we")
    Tooltip(choose_on_map_button, "Automatically inputs UTM zone by copying location from Google maps")

    utm_label = ttk.Label(root, text="UTM zone:", font=('Arial', 8))
    utm_label.grid(row=3, column=0, sticky="e")
    utm_entry = ttk.Entry(root, font=('Arial', 8))
    utm_entry.grid(row=3, column=1, sticky="we")

    hemisphere_label = ttk.Label(root, text="Hemisphere:", font=('Arial', 8))
    hemisphere_label.grid(row=4, column=0, sticky="e")
    hemisphere_combobox = ttk.Combobox(root, values=["inNorthernHemisphere", "inSouthernHemisphere"], state="readonly",
                                       font=('Arial', 8))
    hemisphere_combobox.current(0)
    hemisphere_combobox.grid(row=4, column=1, sticky="we")

    originlat_entry = ttk.Entry(root, font=('Arial', 8))

    originlon_entry = ttk.Entry(root, font=('Arial', 8))


    time1_label = ttk.Label(root, text="Time Period 1(h):", font=('Arial', 8))
    time1_label.grid(row=5, column=0, sticky="e")
    Tooltip(time1_label, "Must be identical to period stated in Aermod input to correctly fetch .plt files for "
                         "conversion")
    time1_entry = ttk.Entry(root, font=('Arial', 8))
    time1_entry.grid(row=5, column=1, sticky="we")

    time2_label = ttk.Label(root, text="Time Period 2(h):", font=('Arial', 8))
    time2_label.grid(row=6, column=0, sticky="e")
    Tooltip(time2_label, "Must be identical to period stated in Aermod input to correctly fetch .plt files for "
                         "conversion")
    time2_entry = ttk.Entry(root, font=('Arial', 8))
    time2_entry.grid(row=6, column=1, sticky="we")

    time3_label = ttk.Label(root, text="Time Period 3(h):", font=('Arial', 8))
    time3_label.grid(row=7, column=0, sticky="e")
    Tooltip(time3_label, "Must be identical to period stated in Aermod input to correctly fetch .plt files for "
                         "conversion")
    time3_entry = ttk.Entry(root, font=('Arial', 8))
    time3_entry.grid(row=7, column=1, sticky="we")

    group_name_label = ttk.Label(root, text="Group Name:", font=('Arial', 8))
    group_name_label.grid(row=8, column=0, sticky="e")
    Tooltip(group_name_label, "Must be identical to name stated in Aermod input to correctly fetch .plt files for "
                              "conversion")
    group_name_entry = ttk.Entry(root, font=('Arial', 8))
    group_name_entry.grid(row=8, column=1, sticky="we")

    min_bin_label = ttk.Label(root, text="Minimum Bin:", font=('Arial', 8))
    min_bin_label.grid(row=9, column=0, sticky="e")
    Tooltip(min_bin_label, "Input format = .5e-9; for defaulting to data range = data ")
    min_bin_entry = ttk.Entry(root, font=('Arial', 8))
    min_bin_entry.grid(row=9, column=1, sticky="we")

    max_bin_label = ttk.Label(root, text="Maximum Bin:", font=('Arial', 8))
    max_bin_label.grid(row=10, column=0, sticky="e")
    Tooltip(max_bin_label, "Input format = .5e-9; for defaulting to data range = data ")
    max_bin_entry = ttk.Entry(root, font=('Arial', 8))
    max_bin_entry.grid(row=10, column=1, sticky="we")

    binningchoice_label = ttk.Label(root, text="Binning Method:", font=('Arial', 8))
    binningchoice_label.grid(row=11, column=0, sticky="e")
    binningchoice_combobox = ttk.Combobox(root, values=["Linear", "Log"], state="readonly", font=('Arial', 8))
    binningchoice_combobox.current(0)
    binningchoice_combobox.grid(row=11, column=1, sticky="we")

    gridcols_label = ttk.Label(root, text="N° Grid Columns:", font=('Arial', 8))
    gridcols_label.grid(row=12, column=0, sticky="e")
    Tooltip(gridcols_label, "Default is 400, increase for larger datasets")
    gridcols_entry = ttk.Entry(root, font=('Arial', 8))
    gridcols_entry.grid(row=12, column=1, sticky="we")

    gridrows_label = ttk.Label(root, text="N° Grid Rows:", font=('Arial', 8))
    gridrows_label.grid(row=13, column=0, sticky="e")
    Tooltip(gridrows_label, "Default is 400, increase for larger datasets")
    gridrows_entry = ttk.Entry(root, font=('Arial', 8))
    gridrows_entry.grid(row=13, column=1, sticky="we")

    smooth_label = ttk.Label(root, text="N° Smoothing Iterations:", font=('Arial', 8))
    smooth_label.grid(row=14, column=0, sticky="e")
    Tooltip(smooth_label, "Default is 1, larger values distort exact locations")
    smooth_entry = ttk.Entry(root, font=('Arial', 8))
    smooth_entry.grid(row=14, column=1, sticky="we")

    contour_label = ttk.Label(root, text="Create Contours:", font=('Arial', 8))
    contour_label.grid(row=15, column=0, sticky="e")
    contour_combobox = ttk.Combobox(root, values=["true", "false"], state="readonly", font=('Arial', 8))
    contour_combobox.current(0)
    contour_combobox.grid(row=15, column=1, sticky="we")

    gradient_label = ttk.Label(root, text="Create Gradient:", font=('Arial', 8))
    gradient_label.grid(row=16, column=0, sticky="e")
    gradient_combobox = ttk.Combobox(root, values=["true", "false"], state="readonly", font=('Arial', 8))
    gradient_combobox.current(0)
    gradient_combobox.grid(row=16, column=1, sticky="we")

    gradientbinningchoice_label = ttk.Label(root, text="Gradient Binning Method:", font=('Arial', 8))
    gradientbinningchoice_label.grid(row=17, column=0, sticky="e")
    gradientbinningchoice_combobox = ttk.Combobox(root, values=["Linear", "Log"], state="readonly", font=('Arial', 8))
    gradientbinningchoice_combobox.current(0)
    gradientbinningchoice_combobox.grid(row=17, column=1, sticky="we")

    compile_button = ttk.Button(root, text="Compile", command=compile_output)
    compile_button.grid(row=19, column=1, sticky="we")
    Tooltip(smooth_label, "Choose folder that includes AERMOD outputs, creates 3 folders with corresponding input "
                          "file in each, also copies the needed aermod .inp, .out and .plt files into folders")

    output_text_label = ttk.Label(root, text="Output Text:")
    output_text_label.grid_forget()
    output_text = tk.Text(root, height=20, width=80)
    output_text.grid_forget()

    root.mainloop()


class AERMODGUI(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.main_window = master
        self.protocol("WM_DELETE_WINDOW", self.on_close_aermodgui)
        self.title("CAIRO for AERMOD ~ “Compile AERMOD input and run output” ")
        self.geometry("600x350")  # Increased window size

        self.stage_labels = ["AERMAP", "AERMOD", "AERPLOT"]
        self.stages = [self.run_aermap, self.run_aermod, self.run_aerplot]
        self.stage_labels_comp = ["AERMAP input compiler", "AERMOD input compiler", "AERPLOT input compiler"]
        self.stages_comp = [app1, app2, app4]

        self.status_labels = []

        self.button_frame = ttk.Frame(self)
        self.button_frame.pack()

        for z, stage_label in enumerate(self.stage_labels):
            button = ttk.Button(self.button_frame, text=stage_label, command=lambda z=z: self.run_stage(z))
            button.grid(row=z, column=1, pady=5)

            status_label = ttk.Label(self.button_frame, text="", foreground="red")
            status_label.grid(row=z, column=2, padx=10)
            self.status_labels.append(status_label)

        self.output_text = tk.Text(self, height=10, width=60, wrap=tk.WORD)
        self.output_text.pack(pady=10)

        for i, (stage_label_comp, stage_comp) in enumerate(zip(self.stage_labels_comp, self.stages_comp)):
            button = ttk.Button(self.button_frame, text=stage_label_comp, command=stage_comp)
            button.grid(row=i, column=0, padx=30, pady=5)

    def run_stage(self, stage_index):
        input_folder = self.choose_input_folder()
        if input_folder:
            if self.stage_labels[stage_index] == "AERMAP":
                executable = config.get("aermap_path", "default_aermap_path.exe")
            elif self.stage_labels[stage_index].startswith("AERMET"):
                executable = os.path.join("C:\\", "AERMOD", "EXE_all", "aermet.exe")
            else:
                executable = config.get("aermod_path", "default_aermod_path.exe")

            print("Executable:", executable)
            print("Input folder:", input_folder)
            os.system(f'"{executable}"')

            if self.stage_labels[stage_index] == "AERMET Stage 1":
                inp_file = "aermet1.inp"
            elif self.stage_labels[stage_index] == "AERMET Stage 2":
                inp_file = "aermet2.inp"
            else:
                inp_file = self.stage_labels[stage_index].lower() + ".inp"

            process = subprocess.Popen([executable, inp_file], cwd=input_folder, shell=True,
                                       stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)

            self.output_text.insert(tk.END, f"Output for {self.stage_labels[stage_index]}:\n")
            for line in process.stdout:
                self.output_text.insert(tk.END, line)
                self.output_text.see(tk.END)
                self.update_idletasks()

            process.wait()

            if self.stage_labels[stage_index] != "AERPLOT":
                if self.stage_labels[stage_index] == "AERMAP":
                    output_file = "aermap.out"
                elif self.stage_labels[stage_index].startswith("AERMET"):
                    output_file = f"aermet_st{stage_index % 2 + 1}.msg"
                else:
                    output_file = "aermod.out"

                if os.path.exists(os.path.join(input_folder, output_file)):
                    print(f"{self.stage_labels[stage_index]} completed successfully.")
                    self.status_labels[stage_index].config(text="✓", foreground="green")
                else:
                    print(f"Error: {self.stage_labels[stage_index]} failed.")
                    self.status_labels[stage_index].config(text="✗", foreground="red")
            else:
                print("AERPLOT completed successfully.")
                self.status_labels[stage_index].config(text="✓", foreground="green")
                self.run_aerplot(input_folder, stage_index)

    def choose_input_folder(self):
        folder_path = filedialog.askdirectory()
        return folder_path

    def run_aermap(self):
        # Placeholder method for running AERMAP
        print("Running AERMAP...")

    def run_aermet_stage1(self):
        # Placeholder method for running AERMET Stage 1
        print("Running AERMET Stage 1...")

    def run_aermet_stage2(self):
        # Placeholder method for running AERMET Stage 2
        print("Running AERMET Stage 2...")

    def run_aermod(self):
        # Placeholder method for running AERMOD
        print("Running AERMOD...")

    def run_aerplot(self, input_folder, stage_index):
        input_folder = filedialog.askdirectory(title="Select the folder containing the necessary files")

        for i in range(1, 4):
            aerplot_folder = os.path.join(input_folder, f"aerplot{i}")
            if os.path.exists(aerplot_folder):
                os.chdir(aerplot_folder)
                subprocess.run(["aerplot"], shell=True)
                print(f"AERPLOT {i} completed successfully.")
            else:
                print(f"Error: Subfolder aerplot{i} not found in {input_folder}.")

    def on_close_aermodgui(self):
        self.destroy()
        if self.main_window:
            self.main_window.destroy()


def main():
    main_window = tk.Tk()
    main_window.title("Main App")
    main_window.withdraw()

    AERMODGUI(main_window)

    # Create buttons to launch each app
    app1_button = tk.Button(main_window, text="Launch App 1", command=app1)
    app1_button.pack()

    app2_button = tk.Button(main_window, text="Launch App 2", command=app2)
    app2_button.pack()

    app3_button = tk.Button(main_window, text="Launch App 3", command=lambda: AERMODGUI(main_window))
    app3_button.pack()

    app4_button = tk.Button(main_window, text="Launch App 2", command=app4)
    app4_button.pack()

    main_window.mainloop()


if __name__ == "__main__":
    main()

