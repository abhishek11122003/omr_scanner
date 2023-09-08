import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from ttkthemes import ThemedStyle
from PIL import Image, ImageTk, ImageDraw, ImageFont
import datetime
import cv2
import numpy as np
import win32api
import re
import os
import time
import glob
import webbrowser

class OMR_Scanner(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("OMR Sheet Checking Application")
        self.geometry("800x660")
        
        # Set the style for a better-looking GUI
        style = ThemedStyle(self)
        style.set_theme("arc")  # Choose from 'clam', 'alt', 'arc', 'classic'

        self.iconbitmap(r"C:\Users\kbnpa\Downloads\Screenshot-2023-07-20-143313.ico")
        
        # Add a header with the application name
        header_frame = tk.Frame(self)
        header_frame.pack(fill="x", padx=10, pady=5)
        header_label = tk.Label(header_frame, text="OMR Sheet Checking Application", font=("Helvetica", 14))
        header_label.pack()
        
        # Create a frame to hold the pages
        self.page_frame = tk.Frame(self)
        self.page_frame.pack(fill="both", expand=True)
        
        self.pages = {}
        
        # Create four different pages
        for PageClass in (INSTRUCTIONS, OMR_CHECKING):
            page_name = PageClass.__name__
            page = PageClass(parent=self.page_frame, controller=self)
            self.pages[page_name] = page
        # Show the first page on startup
        self.show_page("INSTRUCTIONS")
        
    def show_page(self, page_name):
        # Hide all other pages and show the selected page
        for page in self.pages.values():
            page.pack_forget()
        page = self.pages[page_name]
        page.pack(fill="both", expand=True)

class INSTRUCTIONS(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = tk.Label(self, text="USER GUIDE", font=("Helvetica", 15, "bold", "underline"))
        label.pack(pady=5, padx=10)

        # Load the image
        image_path = r"C:\Users\kbnpa\Pictures\Screenshots\Screenshot 2023-07-20 143313.png"
        image = Image.open(image_path)
        image = image.resize((150, 150))  # Resize the image to fit within the GUI

        # Convert the Image object to a PhotoImage object
        self.photo = ImageTk.PhotoImage(image)

        # Create a label to display the image
        self.image_label = tk.Label(self, image=self.photo)
        self.image_label.pack(pady=0)

        text = "NOTE :- For processing OMR properly use\n"
        text1 ="OKEN Scanner App\n"
        text2 ="and use 'Eco' templet.\n" 
        text3 ="1. Step 1:- Select your OMR Sheet image using browse button. \n\n" \
               "2. Step 2:- Select the output folder where you want to save the \n" \
               "           processed OMR Sheet using the browse button.\n\n" \
               "3. Step 3:- Now give the answer key for each subject in            \n" \
               "    (  ) brackets for each question.\n\n" \
               "4. Step 4:- Click on the process image button and wait for the   \n" \
               "  output.\n\n"
        
        frame_text_label = tk.Frame(self)
        self.text_label = tk.Label(frame_text_label, text=text, font=("Helvetica", 10, "bold"))
        self.text_label.grid(row =0, column = 0)

        self.text_label1 = tk.Label(frame_text_label, text= text1, font=("Helvetica", 10, "bold", "underline"), fg="blue", cursor="hand2")
        self.text_label1.grid(row =0, column = 1)
        
        self.text_label2 = tk.Label(frame_text_label, text=text2, font=("Helvetica", 10, "bold"))
        self.text_label2.grid(row =0, column = 2)
        frame_text_label.pack()

        self.text_label3 = tk.Label(self, text=text3, font=("Helvetica", 10, "bold"))
        self.text_label3.pack()

        self.link = "https://play.google.com/store/apps/details?id=com.cambyte.okenscan&hl=en&gl=US&pli=1"
        self.text_label1.bind("<Button-1>", self.on_link_click)

        label = tk.Label(self, text="NOTE: This application is not 100% accurate and only supports a specific type of OMR,\n so it is recommended to verify manually as well!!", font=("Helvetica", 12, "bold"), fg="red")
        label.pack(pady=0, padx=10)

        frame_progress = tk.Frame(self)
        # Create a canvas for the stopwatch animation
        self.canvas = tk.Canvas(frame_progress, width=300, height=20)
        self.canvas.grid(row=4, column=1, padx=10, pady=0)
        # Create a canvas for the custom progress bar
        self.canvas2 = tk.Canvas(frame_progress, width=300, height=10)
        self.canvas2.grid(row=4, column=1, padx=10, pady=0)
        frame_progress.pack()

        # Create a variable to store the start time of the stopwatch
        self.start_time = None
        # Call the update_stopwatch method to start the stopwatch
        self.update_stopwatch()

        # Add "Open and Print Template" button
        self.open_print_button = tk.Button(self, text="Generate OMR", font=("Helvetica", 10), cursor="hand2", command=self.open_file)
        self.open_print_button.pack(pady=5)
        # Add "Next" button
        self.next_button = tk.Button(self, text="PROCEED", font=("Helvetica", 10), cursor="hand2", command=self.go_to_OMR_CHECKING)
        self.next_button.pack(pady=5)

        self.controller = controller

    def on_link_click(self, event):
        webbrowser.open(self.link)

    def go_to_OMR_CHECKING(self):
        self.open_print_button.config(state=tk.DISABLED)
        self.next_button.config(state=tk.DISABLED)
        self.start_stopwatch()
        self.animate_progress(100)
        self.draw_stopwatch(4500)
        self.update_stopwatch()
        self.start_stopwatch()
        self.stop_stopwatch()
        self.reset_progress_bar()
        self.canvas2.delete("all")
        self.controller.show_page("OMR_CHECKING")
        self.open_print_button.config(state=tk.NORMAL)
        self.next_button.config(state=tk.NORMAL)

    def open_file(self):
        file_path = r"C:\Users\kbnpa\Desktop\KV Kpt Class\testpapers\OMR 2.pdf"
        if os.path.exists(file_path):
            self.open_print_button.config(state=tk.DISABLED)
            self.next_button.config(state=tk.DISABLED)
            self.start_stopwatch()
            self.animate_progress(100)
            self.draw_stopwatch(4500)
            self.update_stopwatch()
            self.start_stopwatch()
            self.stop_stopwatch()
            self.reset_progress_bar()
            os.startfile(file_path)
            self.canvas2.delete("all")
            self.open_print_button.config(state=tk.NORMAL)
            self.next_button.config(state=tk.NORMAL)

    # Method to update the stopwatch display (as an animation)
    def update_stopwatch(self):
        if self.start_time:
            elapsed_time = time.time() - self.start_time
            self.draw_stopwatch(elapsed_time)
            if elapsed_time >= 15:  # Stop the stopwatch after 15 seconds and reset it
                self.stop_stopwatch()
                self.canvas.delete("all")  # Clear the canvas
                self.start_time = None
            else:
                self.after(100, self.update_stopwatch)  # Update every 100 milliseconds if not stopped

    # Method to start the stopwatch
    def start_stopwatch(self):
        self.start_time = time.time()
        self.update_stopwatch()  # Start the stopwatch animation

    # Method to stop the stopwatch
    def stop_stopwatch(self):
        self.start_time = None

    # Method to draw the stopwatch on the canvas
    def draw_stopwatch(self, elapsed_time):
        self.canvas.delete("all")  # Clear the canvas
        # Draw a rectangle representing the stopwatch bar
        bar_width = min(int((elapsed_time / 5) * 300), 300)
        self.canvas.create_rectangle(0, 0, bar_width, 18, fill="blue")

    def interpolate_color(self, start_color, end_color, progress):
        r = int(start_color[0] + (end_color[0] - start_color[0]) * progress)
        g = int(start_color[1] + (end_color[1] - start_color[1]) * progress)
        b = int(start_color[2] + (end_color[2] - start_color[2]) * progress)
        return f'#{r:02x}{g:02x}{b:02x}'

    def animate_progress(self, max_value):
        # Define the starting and ending colors for the gradient
        start_color = (255, 0, 0)  # Red in RGB
        end_color = (0, 255, 0)    # Green in RGB
        for i in range(0, max_value + 1):
            # Calculate the current progress value (between 0 and 1)
            progress = i / max_value
            # Interpolate the color based on the progress value
            current_color = self.interpolate_color(start_color, end_color, progress)
            self.canvas2.create_rectangle(0, 0, i * 3, 20, fill=current_color, width=0, tags='progress')
            self.canvas2.update()
            time.sleep(0.045)

    def update_progress(self, value):
        self.canvas2.coords('progress', 0, 0, value * 3, 20)
        self.canvas2.update()

    def reset_progress_bar(self):
        self.canvas2.coords('progress', 0, 0, 0, 20)
        self.canvas2.update()

class OMR_CHECKING(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        # Call the function to delete files from the specified folder
        folder_to_delete_files = r"C:\Users\kbnpa\Desktop\OMR Processing\Processed_OMR"
        self.delete_files_in_folder(folder_to_delete_files)
        Head_frame = tk.Frame(self)
        label = tk.Label(Head_frame, text="                        ", font=("Helvetica", 15))
        label.grid(row=0, column=0, padx=10, pady=0, sticky=tk.W)
        label = tk.Label(Head_frame, text="OMR Processing", font=("Helvetica", 15, "bold", "underline"))
        label.grid(row=0, column=1, padx=10, pady=0, sticky=tk.W)
        # Add "Open and Print Template" button
        self.open_print_button = tk.Button(Head_frame, text="Generate OMR", font=("Helvetica", 10), cursor="hand2", command=self.open_file)
        self.open_print_button.grid(row=0, column=2, padx=10, pady=0, sticky=tk.W)
        # Create the refresh button and place it at the top-right corner
        self.refresh_button = tk.Button(Head_frame, text="Refresh", cursor="hand2", command=self.refresh)
        self.refresh_button.grid(row=0, column=3, padx=10, pady=0, sticky=tk.W)
        self.canvas3 = tk.Canvas(Head_frame, width=72, height=72)
        self.canvas3.grid(row=0, column=0, padx=10, pady=0, sticky=tk.W)
        Head_frame.pack()

        self.image_path = None

        # Input_Image Selection
        Input_Image_frame = tk.Frame(self)
        Input_Image_label = tk.Label(Input_Image_frame, text="Select OMR Image:")
        Input_Image_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.Input_Image_entry = tk.Entry(Input_Image_frame, width=50)
        self.Input_Image_entry.grid(row=0, column=1, padx=10, pady=5)
        self.Input_Image_browse_button = tk.Button(Input_Image_frame, text="Browse", cursor="hand2", command=self.select_image)
        self.Input_Image_browse_button.grid(row=0, column=2, padx=10, pady=5)

        output_folder_label = tk.Label(Input_Image_frame, text="Select Output Folder:")
        output_folder_label.grid(row=1, column=0, padx=10, pady=5)
        self.output_folder_entry = tk.Entry(Input_Image_frame, width=50)
        self.output_folder_entry.grid(row=1, column=1, padx=10, pady=5)
        self.output_folder_browse_button = tk.Button(Input_Image_frame, text="Browse", cursor="hand2", command=self.browse_output_folder)
        self.output_folder_browse_button.grid(row=1, column=2, padx=10, pady=5)
        Input_Image_frame.pack(pady=10)

        # Create a new frame to hold the width, height, and resize frames horizontally
        row_frame = tk.Frame(self)
        width_frame = tk.Frame(row_frame)  # Create a new frame to hold the width label and entry widgets
        self.width_label = tk.Label(width_frame, text="Width (pixels): ", font=("Helvetica", 10))
        self.width_label.pack(side=tk.LEFT)
        self.width_entry = tk.Entry(width_frame, font=("Helvetica", 10), width=5)
        self.width_entry.pack(side=tk.LEFT)
        # Create the "i" button without space
        self.i_button_w = tk.Button(width_frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_width)
        self.i_button_w.pack(side=tk.LEFT, padx=0)  # Set padx=0 to remove horizontal space
        width_frame.pack(side=tk.LEFT, padx=(0, 40))

        height_frame = tk.Frame(row_frame)  # Create a new frame to hold the height label and entry widgets
        self.height_label = tk.Label(height_frame, text="Height (pixels): ", font=("Helvetica", 10))
        self.height_label.pack(side=tk.LEFT)
        self.height_entry = tk.Entry(height_frame, font=("Helvetica", 10), width=5)
        self.height_entry.pack(side=tk.LEFT)
        # Create the "i" button without space
        self.i_button_h = tk.Button(height_frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_height)
        self.i_button_h.pack(side=tk.LEFT, padx=0)
        height_frame.pack(side=tk.LEFT, padx=(0, 40))

        resize_frame = tk.Frame(row_frame)  # Create a new frame to hold the resize label and entry widgets
        label_resize_factor = tk.Label(resize_frame, text="Resize Factor: ", font=("Helvetica", 10))
        label_resize_factor.pack(side=tk.LEFT)
        self.entry_resize_factor = tk.Entry(resize_frame, width=5, font=("Helvetica", 10))
        self.entry_resize_factor.pack(side=tk.LEFT)
        # Create the "i" button without space
        self.i_button_rf = tk.Button(resize_frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_resizefactor)
        self.i_button_rf.pack(side=tk.LEFT, padx=0)
        resize_frame.pack(side=tk.LEFT)
        row_frame.pack(pady=5)

        self.controller = controller
        self.n = 5

        frame = tk.Frame(self)
        physics_label = tk.Label(frame, text="Physics Answer Key:")
        physics_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.physics_entry = tk.Entry(frame, width=60)
        self.physics_entry.grid(row=0, column=1, padx=10, pady=5)
        # Create the "i" button for Physics Answer Key without space
        self.i_button_p = tk.Button(frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_physics)
        self.i_button_p.grid(row=0, column=2, padx=0, pady=5, sticky=tk.W)  # Set padx=0 to remove horizontal space

        chemistry_label = tk.Label(frame, text="Chemistry Answer Key:")
        chemistry_label.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.chemistry_entry = tk.Entry(frame, width=60)
        self.chemistry_entry.grid(row=1, column=1, padx=10, pady=5)
        # Create the "i" button for Physics Answer Key without space
        self.i_button_c = tk.Button(frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_chemistry)
        self.i_button_c.grid(row=1, column=2, padx=0, pady=5, sticky=tk.W)  # Set padx=0 to remove horizontal space

        botany_label = tk.Label(frame, text="Botany Answer Key:")
        botany_label.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.botany_entry = tk.Entry(frame, width=60)
        self.botany_entry.grid(row=2, column=1, padx=10, pady=5)
        # Create the "i" button for Physics Answer Key without space
        self.i_button_b = tk.Button(frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_botany)
        self.i_button_b.grid(row=2, column=2, padx=0, pady=5, sticky=tk.W)  # Set padx=0 to remove horizontal space

        zoology_label = tk.Label(frame, text="Zoology Answer Key:")
        zoology_label.grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        self.zoology_entry = tk.Entry(frame, width=60)
        self.zoology_entry.grid(row=3, column=1, padx=10, pady=5)
        # Create the "i" button for Physics Answer Key without space
        self.i_button_z = tk.Button(frame, text="i", font=("Helvetica", 10, "italic"), width=2, cursor="hand2", command=self.show_information_zoology)
        self.i_button_z.grid(row=3, column=2, padx=0, pady=5, sticky=tk.W)  # Set padx=0 to remove horizontal space
        frame.pack()

        check_frame = tk.Frame(self)
        label_check = tk.Label(check_frame, text="Check Progress: ", font=("Helvetica", 10))
        label_check.pack(side=tk.LEFT)
        self.result_text = tk.Text(check_frame, width=45, height=1, wrap=tk.WORD, state=tk.DISABLED)################
        self.result_text.pack(side=tk.LEFT, padx=5, pady=5)###############################################################
        # Create the button
        open_button_H = tk.Button(check_frame, text="H-Parts", cursor="hand2", command=self.open_folder_H)
        open_button_H.pack(side=tk.LEFT, padx=5)
        # Create the button
        open_button_V = tk.Button(check_frame, text="V-Parts", cursor="hand2", command=self.open_folder_V)
        open_button_V.pack(side=tk.LEFT, padx=5)
        check_frame.pack(pady=10)

        frame_progress = tk.Frame(self)
        # Create a canvas for the stopwatch animation
        self.canvas = tk.Canvas(frame_progress, width=300, height=20)
        self.canvas.grid(row=4, column=1, padx=10, pady=10)
        # Create a canvas for the custom progress bar
        self.canvas2 = tk.Canvas(frame_progress, width=300, height=10)
        self.canvas2.grid(row=4, column=1, padx=10, pady=10)
        frame_progress.pack()

        # Add "NEXT" and "BACK" buttons in the same row
        button_frame = tk.Frame(self)
        self.back_button = tk.Button(button_frame, text="BACK", font=("Helvetica", 10), cursor="hand2", command=self.go_to_step_0)
        self.back_button.pack(side=tk.LEFT, padx=5, pady=5)
        # Create the "Process Image" button and link it to the process_image method
        self.process_button = tk.Button(button_frame, text="Process Image", cursor="hand2", font=("Helvetica", 10), command=self.start_processing)
        self.process_button.pack(side=tk.LEFT, padx=10)
        self.button_print = tk.Button(button_frame, text="Print", cursor="hand2", font=("Helvetica", 10), command=self.open_images_with_button)
        self.button_print.pack(side=tk.LEFT, padx=10)
        button_frame.pack()

        # Create a variable to store the start time of the stopwatch
        self.start_time = None
        # Call the update_stopwatch method to start the stopwatch
        self.update_stopwatch()
        self.controller = controller

        result_frame = tk.Frame(self)
        label_final_result = tk.Label(result_frame, text="RESULT : ", font=("Helvetica", 12, "bold"))
        label_final_result.pack(side=tk.LEFT)
        self.text_box = tk.Text(result_frame, width=35, height=6, wrap=tk.WORD, font=("Helvetica", 12), state=tk.DISABLED)################
        self.text_box.pack(side=tk.LEFT, padx=5, pady=5)###############################################################
        result_frame.pack()

    def open_image_files_in_folder(self, folder_path):
        try:
            if not os.path.exists(folder_path):
                print(f"Folder '{folder_path}' does not exist.")
                return
            file_list = os.listdir(folder_path)
            image_files_found = False
            for file_name in file_list:
                file_path = os.path.join(folder_path, file_name)
                if os.path.isfile(file_path):
                    if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')):
                        image_files_found = True
                        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            if not image_files_found:
                messagebox.showerror("Error", "First process the OMR.")
        except Exception as e:
            print(f"Error occurred: {e}")

    def open_images_with_button(self):
        # Specify the folder path here
        folder_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Processed_OMR"
        self.open_image_files_in_folder(folder_path)

    # Method to draw the circular progress bar on the canvas3
    def draw_progress_bar_cr(self, value):
        self.canvas3.delete("progress_cr")  # Clear the previous progress bar
        center_x = 50  # x-coordinate of the center of the circle
        center_y = 50  # y-coordinate of the center of the circle
        outer_radius = 25   # Outer radius of the ring
        inner_radius = 20   # Inner radius of the ring
        thickness = 5      # Thickness of the ring
        # Calculate the angles for the progress arc
        angle = 360 * value / 100
        # Draw the outer arc (full circle)
        self.canvas3.create_arc(center_x - outer_radius, center_y - outer_radius, center_x + outer_radius, center_y + outer_radius,
                               start=0, extent=360, width=thickness, outline="")
        # Draw the progress arc as an inner ring
        self.canvas3.create_arc(center_x - inner_radius, center_y - inner_radius, center_x + inner_radius, center_y + inner_radius,
                               start=90, extent=-angle, width=thickness, outline="blue", style="arc", tags="ring_progress")

    # Method to animate the progress bar
    def animate_progress_cr(self, max_value):
        for i in range(0, max_value + 1):
            self.draw_progress_bar_cr(i)
            self.canvas3.update()
            time.sleep(0.01)
        # After the animation completes, delete the canvas
        self.canvas3.delete("all")
    
    def update_progress_cr(self, value):
        self.draw_progress_bar_cr(value)
        self.canvas3.update()
    
    def reset_progress_bar_cr(self):
        self.draw_progress_bar_cr(0)  # Reset the progress to 0%
        self.canvas3.update()

    def go_to_step_0(self):
        self.process_button.config(state=tk.DISABLED)
        self.open_print_button.config(state=tk.DISABLED)
        self.i_button_w.config(state=tk.DISABLED)
        self.i_button_h.config(state=tk.DISABLED)
        self.i_button_rf.config(state=tk.DISABLED)
        self.i_button_p.config(state=tk.DISABLED)
        self.i_button_c.config(state=tk.DISABLED)
        self.i_button_b.config(state=tk.DISABLED)
        self.i_button_z.config(state=tk.DISABLED)
        self.back_button.config(state=tk.DISABLED)
        self.button_print.config(state=tk.DISABLED)
        self.Input_Image_browse_button.config(state=tk.DISABLED)
        self.output_folder_browse_button.config(state=tk.DISABLED)
        self.start_stopwatch()
        self.animate_progress(100)
        self.draw_stopwatch(4500)
        self.update_stopwatch()
        self.start_stopwatch()
        self.stop_stopwatch()
        self.reset_progress_bar()
        self.canvas2.delete("all")
        self.controller.show_page("INSTRUCTIONS")
        self.process_button.config(state=tk.NORMAL) 
        self.open_print_button.config(state=tk.NORMAL)
        self.i_button_w.config(state=tk.NORMAL)
        self.i_button_h.config(state=tk.NORMAL)
        self.i_button_rf.config(state=tk.NORMAL)
        self.i_button_p.config(state=tk.NORMAL)
        self.i_button_c.config(state=tk.NORMAL)
        self.i_button_b.config(state=tk.NORMAL)
        self.i_button_z.config(state=tk.NORMAL)
        self.back_button.config(state=tk.NORMAL)
        self.button_print.config(state=tk.NORMAL)
        self.Input_Image_browse_button.config(state=tk.NORMAL)
        self.output_folder_browse_button.config(state=tk.NORMAL)

    def open_file(self):
        file_path = r"C:\Users\kbnpa\Desktop\KV Kpt Class\testpapers\OMR 2.pdf"
        if os.path.exists(file_path):
            self.process_button.config(state=tk.DISABLED)
            self.open_print_button.config(state=tk.DISABLED)
            self.i_button_w.config(state=tk.DISABLED)
            self.i_button_h.config(state=tk.DISABLED)
            self.i_button_rf.config(state=tk.DISABLED)
            self.i_button_p.config(state=tk.DISABLED)
            self.i_button_c.config(state=tk.DISABLED)
            self.i_button_b.config(state=tk.DISABLED)
            self.i_button_z.config(state=tk.DISABLED)
            self.back_button.config(state=tk.DISABLED)
            self.button_print.config(state=tk.DISABLED)
            self.Input_Image_browse_button.config(state=tk.DISABLED)
            self.output_folder_browse_button.config(state=tk.DISABLED)
            self.start_stopwatch()
            self.animate_progress(100)
            self.draw_stopwatch(4500)
            self.update_stopwatch()
            self.start_stopwatch()
            self.stop_stopwatch()
            self.reset_progress_bar()
            os.startfile(file_path)
            self.canvas2.delete("all")
            self.process_button.config(state=tk.NORMAL) 
            self.open_print_button.config(state=tk.NORMAL)
            self.i_button_w.config(state=tk.NORMAL)
            self.i_button_h.config(state=tk.NORMAL)
            self.i_button_rf.config(state=tk.NORMAL)
            self.i_button_p.config(state=tk.NORMAL)
            self.i_button_c.config(state=tk.NORMAL)
            self.i_button_b.config(state=tk.NORMAL)
            self.i_button_z.config(state=tk.NORMAL)
            self.back_button.config(state=tk.NORMAL)
            self.button_print.config(state=tk.NORMAL)
            self.Input_Image_browse_button.config(state=tk.NORMAL)
            self.output_folder_browse_button.config(state=tk.NORMAL)

    def open_folder_H(self):
        folder_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Horizontal_Split"  # Replace this with your desired folder path
        os.startfile(folder_path)

    def open_folder_V(self):
        folder_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Vertical_Split"  # Replace this with your desired folder path
        os.startfile(folder_path)

    # Method to start the processing (and stopwatch)
    def start_processing(self):
        #self.start_stopwatch()  # Start the stopwatch
        self.process_image()  # Call the existing process_image method

    # Method to update the stopwatch display (as an animation)
    def update_stopwatch(self):
        if self.start_time:
            elapsed_time = time.time() - self.start_time
            self.draw_stopwatch(elapsed_time)
            if elapsed_time >= 15:  # Stop the stopwatch after 15 seconds and reset it
                self.stop_stopwatch()
                self.canvas.delete("all")  # Clear the canvas
                self.start_time = None
            else:
                self.after(100, self.update_stopwatch)  # Update every 100 milliseconds if not stopped

    # Method to start the stopwatch
    def start_stopwatch(self):
        self.start_time = time.time()
        self.update_stopwatch()  # Start the stopwatch animation

    # Method to stop the stopwatch
    def stop_stopwatch(self):
        self.start_time = None

    # Method to draw the stopwatch on the canvas
    def draw_stopwatch(self, elapsed_time):
        self.canvas.delete("all")  # Clear the canvas
        # Draw a rectangle representing the stopwatch bar
        bar_width = min(int((elapsed_time / 5) * 300), 300)
        self.canvas.create_rectangle(0, 0, bar_width, 18, fill="blue")

    def refresh(self):
        # Call the function to delete files from the specified folder
        folder_to_delete_files = r"C:\Users\kbnpa\Desktop\OMR Processing\Processed_OMR"
        self.delete_files_in_folder(folder_to_delete_files)
        self.process_button.config(state=tk.DISABLED)
        self.open_print_button.config(state=tk.DISABLED)
        self.i_button_w.config(state=tk.DISABLED)
        self.i_button_h.config(state=tk.DISABLED)
        self.i_button_rf.config(state=tk.DISABLED)
        self.i_button_p.config(state=tk.DISABLED)
        self.i_button_c.config(state=tk.DISABLED)
        self.i_button_b.config(state=tk.DISABLED)
        self.i_button_z.config(state=tk.DISABLED)
        self.back_button.config(state=tk.DISABLED)
        self.button_print.config(state=tk.DISABLED)
        self.Input_Image_browse_button.config(state=tk.DISABLED)
        self.output_folder_browse_button.config(state=tk.DISABLED)
        self.animate_progress_cr(100)  # Animate the progress from 0% to 100%
        self.update_progress_cr(0)    # Update the progress to 0%
        # Change state to normal to clear the Text widgets
        self.text_box.config(state=tk.NORMAL)
        self.result_text.config(state=tk.NORMAL)
        self.output_folder_entry.config(state=tk.NORMAL)
        self.Input_Image_entry.config(state=tk.NORMAL)
        self.width_entry.config(state=tk.NORMAL)
        self.height_entry.config(state=tk.NORMAL)
        self.entry_resize_factor.config(state=tk.NORMAL)
        self.physics_entry.config(state=tk.NORMAL)
        self.chemistry_entry.config(state=tk.NORMAL)
        self.botany_entry.config(state=tk.NORMAL)
        self.zoology_entry.config(state=tk.NORMAL)
        self.animate_progress_cr(100)  # Animate the progress from 0% to 100%
        self.update_progress_cr(0)    # Update the progress to 0%
        self.reset_progress_bar()
        # Clear the Text widgets
        self.text_box.delete("1.0", tk.END)
        self.result_text.delete("1.0", tk.END)
        self.output_folder_entry.delete(0, tk.END)
        self.Input_Image_entry.delete(0, tk.END)
        self.width_entry.delete(0,tk.END)
        self.height_entry.delete(0,tk.END)
        self.entry_resize_factor.delete(0,tk.END)
        self.physics_entry.delete(0,tk.END)
        self.chemistry_entry.delete(0,tk.END)
        self.botany_entry.delete(0,tk.END)
        self.zoology_entry.delete(0,tk.END)
        # Set the Text widgets back to disabled
        self.text_box.config(state=tk.DISABLED)
        self.result_text.config(state=tk.DISABLED)
        self.animate_progress_cr(100)  # Animate the progress from 0% to 100%
        self.update_progress_cr(0)    # Update the progress to 0%
        # Enable the process_button
        self.process_button.config(state=tk.NORMAL)
        self.open_print_button.config(state=tk.NORMAL)
        self.i_button_w.config(state=tk.NORMAL)
        self.i_button_h.config(state=tk.NORMAL)
        self.i_button_rf.config(state=tk.NORMAL)
        self.i_button_p.config(state=tk.NORMAL)
        self.i_button_c.config(state=tk.NORMAL)
        self.i_button_b.config(state=tk.NORMAL)
        self.i_button_z.config(state=tk.NORMAL)
        self.back_button.config(state=tk.NORMAL)
        self.button_print.config(state=tk.NORMAL)
        self.Input_Image_browse_button.config(state=tk.NORMAL)
        self.output_folder_browse_button.config(state=tk.NORMAL)

    def select_image(self):
        self.image_path = filedialog.askopenfilename(initialdir="/", title="Select Image",
                                                    filetypes=(("JPEG files", "*.jpg"), ("all files", "*.*")))
        if self.image_path:
            self.Input_Image_entry.delete(0, tk.END)  # Clear the entry widget
            self.Input_Image_entry.insert(0, self.image_path)  # Set the entry text to the selected image path

    def process_image(self):
        self.process_button.config(state=tk.DISABLED)
        self.open_print_button.config(state=tk.DISABLED)
        self.i_button_w.config(state=tk.DISABLED)
        self.i_button_h.config(state=tk.DISABLED)
        self.i_button_rf.config(state=tk.DISABLED)
        self.i_button_p.config(state=tk.DISABLED)
        self.i_button_c.config(state=tk.DISABLED)
        self.i_button_b.config(state=tk.DISABLED)
        self.i_button_z.config(state=tk.DISABLED)
        self.back_button.config(state=tk.DISABLED)
        self.button_print.config(state=tk.DISABLED)
        self.Input_Image_browse_button.config(state=tk.DISABLED)
        self.output_folder_browse_button.config(state=tk.DISABLED)

        self.animate_progress(100)
        # Simulating the image processing by waiting for 5 seconds
        for i in range(101):
            time.sleep(0.01)  # Simulating image processing time (you can replace this with the actual image processing code)
            self.update_progress(i)
        # Call the function to delete files from the specified folder
        folder_to_delete_files = r"C:\Users\kbnpa\Desktop\OMR Processing\Vertical_Split"
        self.delete_files_in_folder(folder_to_delete_files)
        
        if self.image_path:
            num_parts_str_V = 4
            width_str = self.width_entry.get()
            height_str = self.height_entry.get()
            resize_factor_str = self.entry_resize_factor.get()

            if width_str.isdigit() and height_str.isdigit():
                num_parts_V = int(num_parts_str_V)
                widthImg = int(width_str)
                heightImg = int(height_str)
                output_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Vertical_Split"
                num_parts_H = 10

                if resize_factor_str:
                    resize_factor = float(resize_factor_str)
                    def stackImages(imgArray,scale,lables=[]):
                        rows = len(imgArray)
                        cols = len(imgArray[0])
                        rowsAvailable = isinstance(imgArray[0], list)
                        width = imgArray[0][0].shape[1]
                        height = imgArray[0][0].shape[0]
                        if rowsAvailable:
                            for x in range ( 0, rows):
                                for y in range(0, cols):
                                    imgArray[x][y] = cv2.resize(imgArray[x][y], (0, 0), None, scale, scale)
                                    if len(imgArray[x][y].shape) == 2: imgArray[x][y]= cv2.cvtColor( imgArray[x][y], cv2.COLOR_GRAY2BGR)
                            imageBlank = np.zeros((height, width, 3), np.uint8)
                            hor = [imageBlank]*rows
                            hor_con = [imageBlank]*rows
                            for x in range(0, rows):
                                hor[x] = np.hstack(imgArray[x])
                                hor_con[x] = np.concatenate(imgArray[x])
                            ver = np.vstack(hor)
                            ver_con = np.concatenate(hor)
                        else:
                            for x in range(0, rows):
                                imgArray[x] = cv2.resize(imgArray[x], (0, 0), None, scale, scale)
                                if len(imgArray[x].shape) == 2: imgArray[x] = cv2.cvtColor(imgArray[x], cv2.COLOR_GRAY2BGR)
                            hor= np.hstack(imgArray)
                            hor_con= np.concatenate(imgArray)
                            ver = hor
                        if len(lables) != 0:
                            eachImgWidth= int(ver.shape[1] / cols)
                            eachImgHeight = int(ver.shape[0] / rows)
                            #print(eachImgHeight)
                            for d in range(0, rows):
                                for c in range (0,cols):
                                    cv2.rectangle(ver,(c*eachImgWidth,eachImgHeight*d),(c*eachImgWidth+len(lables[d][c])*13+27,30+eachImgHeight*d),(255,255,255),cv2.FILLED)
                                    cv2.putText(ver,lables[d][c],(eachImgWidth*c+10,eachImgHeight*d+20),cv2.FONT_HERSHEY_COMPLEX,0.7,(255,0,255),2)
                        return ver

                    def rectContour(contours):
                        rectCon = []
                        for i in contours:
                            area = cv2.contourArea(i)
                            if area > 50:
                                peri = cv2.arcLength(i, True)
                                approx = cv2.approxPolyDP(i, 0.04 * peri, True)
                                if len(approx) == 4:
                                    rectCon.append(i)
                        rectCon = sorted(rectCon, key=cv2.contourArea, reverse=True)
                        return rectCon

                    def getCornerPoints(cont):
                        peri = cv2.arcLength(cont, True)
                        approx = cv2.approxPolyDP(cont, 0.04 * peri, True)
                        return approx

                    def reorder(myPoints):
                        myPoints = myPoints.reshape((4, 2))
                        myPointsNew = np.zeros((4, 1, 2), np.int32)
                        add = myPoints.sum(1)
                        myPointsNew[0] = myPoints[np.argmin(add)]  # [0, 0]
                        myPointsNew[3] = myPoints[np.argmax(add)]  # [w, h]
                        diff = np.diff(myPoints, axis=1)
                        myPointsNew[1] = myPoints[np.argmin(diff)]  # [w, 0]
                        myPointsNew[2] = myPoints[np.argmax(diff)]  # [h, 0]
                        return myPointsNew

                    def split_image_vertical(img, num_parts_V, output_path):
                        height, width, _ = img.shape
                        part_width = width // num_parts_V
                        for i in range(num_parts_V):
                            start_x = i * part_width
                            end_x = (i + 1) * part_width
                            part = img[:, start_x:end_x]
                            current_time = time.strftime("%Y%m%d_%H%M%S")
                            filename = os.path.join(output_path, f"vertical_part{i + 1}_{current_time}.jpg")        
                            cv2.imwrite(filename, part)
                            print(filename)

                    # Delete previously saved images in the output directory
                    files = glob.glob(output_path + "/*.jpg")
                    for file in files:
                        os.remove(file)

                    # PRE-PROCESSING
                    path = self.image_path
                    img = cv2.imread(path)
                    img = cv2.resize(img, (widthImg, heightImg), interpolation=cv2.INTER_NEAREST)#####
                    imgContours = img.copy()
                    imgBiggestContours = img.copy()
                    imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    imgBlur = cv2.GaussianBlur(imgGray, (7, 7), 1)#####
                    imgCanny = cv2.Canny(imgBlur, threshold1=20, threshold2=100)#####

                    # FINDING ALL CONTOURS
                    contours, hierarchy = cv2.findContours(imgCanny, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
                    cv2.drawContours(imgContours, contours, -1, (0, 255, 0), 6)
                    # FIND RECTANGLES
                    rectCon = rectContour(contours)
                    biggestContour = getCornerPoints(rectCon[0])

                    if biggestContour.size != 0:
                        cv2.drawContours(imgBiggestContours, biggestContour, -1, (0, 0, 255), 20)
                        biggestContour = reorder(biggestContour)
                        pt1 = np.float32(biggestContour)
                        pt2 = np.float32([[0, 0], [widthImg, 0], [0, heightImg], [widthImg, heightImg]])
                        matrix1 = cv2.getPerspectiveTransform(pt1, pt2)
                        imgwarpColored = cv2.warpPerspective(img, matrix1, (widthImg, heightImg))

                        split_image_vertical(imgwarpColored, num_parts_V, output_path)

                    imgBlank = np.zeros_like(img)
                    imageArry = [
                        [img, imgGray, imgBlur, imgCanny],
                        [imgContours, imgBiggestContours, imgwarpColored, imgBlank]
                    ]

                    # Displaying the processed image
                    imgStacked = stackImages(imageArry, 0.3)
                    #cv2.imshow("Processed Image", imgStacked)

                    # Get input values from the GUI
                    folder_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Vertical_Split"
                    output_path = r"C:\Users\kbnpa\Desktop\OMR Processing\Horizontal_Split"

                    def split_image_horizontal(img_H, num_parts_H, output_path):
                        height, width, _ = img_H.shape
                        part_height = height // num_parts_H
                        for i in range(num_parts_H):
                            start_y = i * part_height
                            end_y = (i + 1) * part_height
                            part = img_H[start_y:end_y, :]
                            current_time = time.strftime("%Y%m%d_%H%M%S")
                            filename = os.path.join(output_path, f"horizontal_part{i + 1}_{current_time}.jpg")
                            cv2.imwrite(filename, part)
                            print(filename)

                    # Get a list of all image files in the folder
                    image_files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

                    # Remove the files present in the output directory before running the code
                    files = glob.glob(output_path + "/*.jpg")
                    for file in files:
                        os.remove(file)

                    # Process each image file
                    for image_file in image_files:
                        image_path = os.path.join(folder_path, image_file)
                        img_H = cv2.imread(image_path)

                        # PRE-PROCESSING
                        # Increase the size of the image
                        img_H = cv2.resize(img_H, (0, 0), fx=resize_factor, fy=resize_factor)

                        # Perform the rest of the image processing steps
                        img_HContours = img_H.copy()
                        img_HBiggestContours = img_H.copy()
                        img_HGray = cv2.cvtColor(img_H, cv2.COLOR_BGR2GRAY)
                        img_HBlur = cv2.GaussianBlur(img_HGray, (5, 5), 1)
                        img_HCanny = cv2.Canny(img_HBlur, 10, 50)

                        # FINDING ALL CONTOURS
                        contours, hierarchy = cv2.findContours(img_HCanny, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
                        cv2.drawContours(img_HContours, contours, -1, (0, 255, 0), 6)

                        # FIND RECTANGLES
                        rectCon = rectContour(contours)
                        biggestContour = getCornerPoints(rectCon[0])

                        if biggestContour.size != 0:
                            cv2.drawContours(img_HBiggestContours, biggestContour, -1, (0, 0, 255), 20)
                            # Perform perspective transformation
                            biggestContour = reorder(biggestContour)
                            pt1 = np.float32(biggestContour)
                            pt2 = np.float32([[0, 0], [img_H.shape[1], 0], [0, img_H.shape[0]], [img_H.shape[1], img_H.shape[0]]])
                            matrix1 = cv2.getPerspectiveTransform(pt1, pt2)
                            img_HwarpColored = cv2.warpPerspective(img_H, matrix1, (img_H.shape[1], img_H.shape[0]))

                            # Perform horizontal splitting
                            split_image_horizontal(img_HwarpColored, num_parts_H, output_path)

                        img_HBlank = np.zeros_like(img_H)
                        imageArray = [
                            [img_H, img_HGray, img_HBlur, img_HCanny],
                            [img_HContours, img_HBiggestContours, img_HwarpColored, img_HBlank]
                        ]

                        img_HStacked_H = stackImages(imageArray, 0.3)
                        #cv2.imshow("Stacked Images", img_HStacked_H)
                        #cv2.imshow("BiggestContours", img_HBiggestContours)
                        cv2.waitKey(3000)

                    # Close any open windows
                    cv2.destroyAllWindows()
                    self.generate_combined_list()
                else:
                    messagebox.showerror("Error", "Invalid resize factor !")
                    time.sleep(0.03)
                    self.canvas2.delete("all")
            else:
                messagebox.showerror("Error", "Invalid image width or image height !")
                time.sleep(0.03)
                self.canvas2.delete("all")
        else:
            messagebox.showerror("Error", "Please select an image first !")
            time.sleep(0.03)
            self.canvas2.delete("all")

        # Reset the progress bar to its initial state
        self.update_progress(0)
        time.sleep(0)

        # Enable the button again after the output is shown
        self.process_button.config(state=tk.NORMAL) 
        self.open_print_button.config(state=tk.NORMAL)
        self.i_button_w.config(state=tk.NORMAL)
        self.i_button_h.config(state=tk.NORMAL)
        self.i_button_rf.config(state=tk.NORMAL)
        self.i_button_p.config(state=tk.NORMAL)
        self.i_button_c.config(state=tk.NORMAL)
        self.i_button_b.config(state=tk.NORMAL)
        self.i_button_z.config(state=tk.NORMAL)
        self.back_button.config(state=tk.NORMAL)
        self.button_print.config(state=tk.NORMAL)
        self.Input_Image_browse_button.config(state=tk.NORMAL)
        self.output_folder_browse_button.config(state=tk.NORMAL)       

    def browse_output_folder(self):
        output_folder_selected = filedialog.askdirectory()
        self.output_folder_entry.delete(0, tk.END)
        self.output_folder_entry.insert(0, output_folder_selected)

    def show_information_width(self):
        info_message = "Give apropriate width for the OMR Sheet image.\n ( Should a 3 or 4 digit natural number )\n Try using value = 1300.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_height(self):
        info_message = "Give apropriate height for the OMR Sheet image.\n ( Should a 3 or 4 digit natural number )\n Try using value = 2200.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_resizefactor(self):
        info_message = "Give apropriate resize factor for the OMR Sheet image.\n ( It can be a natural or decimal number )\n Try using value = 2.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_physics(self):
        info_message = "Give answer key for physics in parenthesis for each questions.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_chemistry(self):
        info_message = "Give answer key for chemistry in parenthesis for each questions.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_botany(self):
        info_message = "Give answer key for botany in parenthesis for each questions.\n" \
        
        messagebox.showinfo("Information", info_message)

    def show_information_zoology(self):
        info_message = "Give answer key for zoology in parenthesis for each questions.\n" \
        
        messagebox.showinfo("Information", info_message)

    def interpolate_color(self, start_color, end_color, progress):
        r = int(start_color[0] + (end_color[0] - start_color[0]) * progress)
        g = int(start_color[1] + (end_color[1] - start_color[1]) * progress)
        b = int(start_color[2] + (end_color[2] - start_color[2]) * progress)
        return f'#{r:02x}{g:02x}{b:02x}'

    def animate_progress(self, max_value):
        # Define the starting and ending colors for the gradient
        start_color = (255, 0, 0)  # Red in RGB
        end_color = (0, 255, 0)    # Green in RGB
        for i in range(0, max_value + 1):
            # Calculate the current progress value (between 0 and 1)
            progress = i / max_value
            # Interpolate the color based on the progress value
            current_color = self.interpolate_color(start_color, end_color, progress)
            self.canvas2.create_rectangle(0, 0, i * 3, 20, fill=current_color, width=0, tags='progress')
            self.canvas2.update()
            time.sleep(0.045)

    def update_progress(self, value):
        self.canvas2.coords('progress', 0, 0, value * 3, 20)
        self.canvas2.update()

    def reset_progress_bar(self):
        self.canvas2.coords('progress', 0, 0, 0, 20)
        self.canvas2.update()

    def extract_numbers_in_parentheses(self, text):
        pattern = r'\((\d+)\)'
        numbers_in_parentheses = re.findall(pattern, text)
        return "".join(numbers_in_parentheses)

    def split_into_lists(self, values):
        def char_to_int(char):
            if char.isdigit():
                return int(char)
            elif 'A' <= char <= 'Z':
                return ord(char) - ord('A') + 1
            else:
                raise ValueError("Invalid character in the answer key")
        result = [list(map(char_to_int, values[i:i+self.n])) for i in range(0, len(values), self.n)]
        return result

    def generate_combined_list(self):
        physics_ans_key = self.physics_entry.get()
        chemistry_ans_key = self.chemistry_entry.get()
        botany_ans_key = self.botany_entry.get()
        zoology_ans_key = self.zoology_entry.get()

        result_1 = self.extract_numbers_in_parentheses(physics_ans_key)
        result_2 = self.extract_numbers_in_parentheses(chemistry_ans_key)
        result_3 = self.extract_numbers_in_parentheses(botany_ans_key)
        result_4 = self.extract_numbers_in_parentheses(zoology_ans_key)

        if not (result_1 and result_2 and result_3 and result_4):
            messagebox.showerror("Error", "Enter answer of each question with in (  ) brackets.\n For more detail check User Guide.")
            time.sleep(0.03)
            self.canvas2.delete("all")
            return

        result_list_1 = self.split_into_lists(result_1)
        result_list_2 = self.split_into_lists(result_2)
        result_list_3 = self.split_into_lists(result_3)
        result_list_4 = self.split_into_lists(result_4)

        combined_list = []
        for i in range(min(len(result_list_1), len(result_list_2), len(result_list_3), len(result_list_4))):
            combined_list.extend([result_list_1[i], result_list_2[i], result_list_3[i], result_list_4[i]])

        total_sublists_combined = len(combined_list)

        def cut_and_paste_sublists(lst):
            num_sublists = len(lst)
            if num_sublists >= 4:
                cut_sublists = lst[-4:]
                remaining_sublists = lst[:-4]
                new_list = cut_sublists + remaining_sublists
                return new_list
            else:
                messagebox.showerror("Error", "The list has fewer than 4 sublists. Cannot perform the operation.")
                return lst

        # Example usage:
        original_list = combined_list
        result_list = cut_and_paste_sublists(original_list)
        self.formatted_combined_list = "[\n{}\n]".format(",\n".join("[{}]".format(", ".join(map(str, sublist))) for sublist in result_list))
        #print(result_list)

        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, f"Total number of sublists in combined list: {total_sublists_combined}")
        self.result_text.config(state=tk.DISABLED)

        self.process_omr()

    widthImg = 400
    heightImg = 400
    questions = 5          # NOTE: Put, questions = no. of questions in each block
    choices = 5
    subjects = 4
    space = 50
    threshold_value = 65   # NOTE: Adjust the threshold_value according to the input image

    def process_answers(self,answers):
        ans = []
        for i in range(0, self.questions):
            if answers[i] == 5:
                ans.append(4)
            if answers[i] == 4:
                ans.append(3)
            if answers[i] == 3:
                ans.append(2)
            if answers[i] == 2:
                ans.append(1)
            if answers[i] == 1:
                ans.append(0)
        return ans
    
    ## TO STACK ALL THE IMAGES IN ONE WINDOW
    def stackImages(self,imgArray, scale, labels=[]):
        rows = len(imgArray)
        cols = len(imgArray[0])
        rowsAvailable = isinstance(imgArray[0], list)
        width = imgArray[0][0].shape[1]
        height = imgArray[0][0].shape[0]
        if rowsAvailable:
            for x in range(0, rows):
                for y in range(0, cols):
                    imgArray[x][y] = cv2.resize(imgArray[x][y], (0, 0), None, scale, scale)
                    if len(imgArray[x][y].shape) == 2:
                        imgArray[x][y] = cv2.cvtColor(imgArray[x][y], cv2.COLOR_GRAY2BGR)
            imageBlank = np.zeros((height, width, 3), np.uint8)
            hor = [imageBlank] * rows
            hor_con = [imageBlank] * rows
            for x in range(0, rows):
                hor[x] = np.hstack(imgArray[x])
                hor_con[x] = np.concatenate(imgArray[x])
            ver = np.vstack(hor)
            ver_con = np.concatenate(hor)
        else:
            for x in range(0, rows):
                imgArray[x] = cv2.resize(imgArray[x], (0, 0), None, scale, scale)
                if len(imgArray[x].shape) == 2:
                    imgArray[x] = cv2.cvtColor(imgArray[x], cv2.COLOR_GRAY2BGR)
            hor = np.hstack(imgArray)
            hor_con = np.concatenate(imgArray)
            ver = hor
        if len(labels) != 0:
            eachImgWidth = int(ver.shape[1] / cols)
            eachImgHeight = int(ver.shape[0] / rows)
            for d in range(0, rows):
                for c in range(0, cols):
                    cv2.rectangle(ver, (c * eachImgWidth, eachImgHeight * d),
                                  (c * eachImgWidth + len(labels[d][c]) * 13 + 27, 30 + eachImgHeight * d),
                                  (255, 255, 255), cv2.FILLED)
                    cv2.putText(ver, labels[d][c], (eachImgWidth * c + 10, eachImgHeight * d + 20),
                                cv2.FONT_HERSHEY_COMPLEX, 0.7, (255, 0, 255), 2)
        return ver
    
    def stack_images_col_wise(self, images, n, label_text_image):
        # Calculate the total number of rows required
        rows = int(np.ceil(len(images) / n))
        # Calculate the common size for all images
        max_height = max(image.shape[0] for image in images)
        total_width = sum(image.shape[1] for image in images[:n])  # Width of the first row
        # Create a blank canvas to stack the images
        stacked_image = np.zeros((max_height * rows, total_width, 3), dtype=np.uint8)
        # Place the images onto the canvas
        x_offset, y_offset = 0, 0
        for i, image in enumerate(images):
            h, w, _ = image.shape
            stacked_image[y_offset:y_offset + h, x_offset:x_offset + w] = image
            x_offset += w
            if (i + 1) % n == 0:
                y_offset += h
                x_offset = 0
        # Cut the first row of pixels
        first_row_cut = stacked_image[:max_height]
        # Continue stacking the rest of the images without the first row
        rest_stacked_image = stacked_image[max_height:]
        # Place the cut first row just after the last row
        stacked_image = np.vstack((rest_stacked_image, first_row_cut))
        # Define the width and height of the text box
        text_box_width = 700
        text_box_height = 550
        blank_width = 801  # You can adjust this value as per your requirements
        # Create a white blank canvas with the new width
        blank_canvas = np.ones((stacked_image.shape[0], blank_width, 3), dtype=np.uint8) * 255
        # Convert the blank canvas to a Pillow Image object
        blank_canvas_pil = Image.fromarray(blank_canvas)
        # Add the text overlay to the blank canvas using Pillow
        draw = ImageDraw.Draw(blank_canvas_pil)
        font = ImageFont.truetype("arial.ttf", 50)  # You can use a different font if needed
        # Calculate the text position within the box
        text_x = 20
        text_y = 70
        # Draw the rectangle for the text box
        draw.rectangle([0, 0, text_box_width - 1, text_box_height - 1], outline=(0, 0, 0), width=2)
        # Draw the text inside the box
        draw.text((text_x, text_y), label_text_image, fill=(0, 0, 0), font=font)
        # Convert the Pillow Image back to a NumPy array
        blank_canvas_with_text = np.array(blank_canvas_pil)
        # Concatenate the blank canvas with text to the left side of the stacked image
        final_stacked_image = np.hstack((blank_canvas_with_text, stacked_image))
        # Create a blank canvas for the right side
        right_blank_width = 100  # You can adjust this value as per your requirements
        right_blank_canvas = np.ones((final_stacked_image.shape[0], right_blank_width, 3), dtype=np.uint8) * 255
        # Concatenate the right blank canvas to the right side of the final_stacked_image
        final_stacked_image_with_right_blank = np.hstack((final_stacked_image, right_blank_canvas))
        # Create a blank canvas for the top
        top_blank_height = 100  # You can adjust this value as per your requirements
        top_blank_width = final_stacked_image_with_right_blank.shape[1]
        top_blank_canvas = np.ones((top_blank_height, top_blank_width, 3), dtype=np.uint8) * 255
        # Concatenate the top blank canvas above the final_stacked_image_with_right_blank
        final_stacked_image_with_blank_top = np.vstack((top_blank_canvas, final_stacked_image_with_right_blank))
        # Create a blank canvas for the bottom
        bottom_blank_height = 100  # You can adjust this value as per your requirements
        bottom_blank_width = final_stacked_image_with_blank_top.shape[1]
        bottom_blank_canvas = np.ones((bottom_blank_height, bottom_blank_width, 3), dtype=np.uint8) * 255
        # Concatenate the bottom blank canvas below the final_stacked_image_with_blank_top
        final_stacked_image_with_blank_top_and_bottom = np.vstack((final_stacked_image_with_blank_top, bottom_blank_canvas))
        # Create a blank canvas for the left
        left_blank_width = 100  # You can adjust this value as per your requirements
        total_width_with_left_blank = final_stacked_image_with_blank_top_and_bottom.shape[1] + left_blank_width
        left_blank_canvas = np.ones((final_stacked_image_with_blank_top_and_bottom.shape[0], left_blank_width, 3), dtype=np.uint8) * 255
        # Concatenate the left blank canvas to the left of the final_stacked_image_with_blank_top_and_bottom
        final_stacked_image_with_blank_sides = np.hstack((left_blank_canvas, final_stacked_image_with_blank_top_and_bottom))
        return final_stacked_image_with_blank_sides
    
#    def stack_images_col_wise_1(self, images, n):
#        # Calculate the total number of rows required
#        rows = int(np.ceil(len(images) / n))
#        # Calculate the common size for all images
#        max_height = max(image.shape[0] for image in images)
#        total_width = sum(image.shape[1] for image in images[:n])  # Width of the first row
#        # Create a blank canvas to stack the images
#        stacked_image = np.zeros((max_height * rows, total_width, 3), dtype=np.uint8)
#        # Place the images onto the canvas
#        x_offset, y_offset = 0, 0
#        for i, image in enumerate(images):
#            h, w, _ = image.shape
#            stacked_image[y_offset:y_offset + h, x_offset:x_offset + w] = image
#            x_offset += w
#            if (i + 1) % n == 0:
#                y_offset += h
#                x_offset = 0
#        # Cut the first row of pixels
#        first_row_cut = stacked_image[:max_height]
#        # Continue stacking the rest of the images without the first row
#        rest_stacked_image = stacked_image[max_height:]
#        # Place the cut first row just after the last row
#        stacked_image = np.vstack((rest_stacked_image, first_row_cut))
#        return stacked_image################################################################################

    def preprocessImage(self,img):
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # Apply adaptive thresholding
        thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
        # Apply morphological operations (closing) to fill gaps in contours
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
        return closed
    
    def rectContour(self,contours):
        rectCon = []
        for i in contours:
            area = cv2.contourArea(i)
            if area > 50:
                peri = cv2.arcLength(i, True)
                approx = cv2.approxPolyDP(i, 0.04 * peri, True)
                if len(approx) == 4:
                    rectCon.append(i)
        rectCon = sorted(rectCon, key=cv2.contourArea, reverse=True)
        return rectCon
    
    def getCornerPoints(self,cont):
        peri = cv2.arcLength(cont, True)
        approx = cv2.approxPolyDP(cont, 0.04 * peri, True)
        return approx
    
    def reorder(self,myPoints):
        myPoints = myPoints.reshape((4, 2))
        myPointsNew = np.zeros((4, 1, 2), np.int32)
        add = myPoints.sum(1)
        myPointsNew[0] = myPoints[np.argmin(add)]
        myPointsNew[3] = myPoints[np.argmax(add)]
        diff = np.diff(myPoints, axis=1)
        myPointsNew[1] = myPoints[np.argmin(diff)]
        myPointsNew[2] = myPoints[np.argmax(diff)]
        return myPointsNew
    
    def splitBoxes(self,img):
        rows = np.vsplit(img, self.choices)
        boxes = []
        for r in rows:
            cols = np.hsplit(r, self.questions)
            for box in cols:
                boxes.append(box)
        return boxes
    
    def showAnswers(self, img, myIndex, grading, ans, questions, choices):
        secW = int(img.shape[1] / questions)
        secH = int(img.shape[0] / choices)
        for x in range(0, questions):
            myAns = myIndex[x]
            cX = (myAns * secW) + secW // 2
            cY = (x * secH) + secH // 2
            if grading[x] == 4:  # Correct answer
                myColor = (0, 255, 0)  # Green
            else:
                if myAns == 4:  # Wrong answer and marked 5th choice
                    myColor = (0, 255, 255)  # Yellow
                else:
                    myColor = (0, 0, 255)  # Red
                correctAns = ans[x]
                cv2.circle(img, ((correctAns * secW) + secW // 2, (x * secH) + secH // 2), 15, (255, 255, 0), cv2.FILLED)
            cv2.circle(img, (cX, cY), 30, myColor, cv2.FILLED)
        return img
    
    def save_stacked_images(self,stacked_images, output_folder):
        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_filename = "stacked_images_" + current_time + ".jpg"
        output_path = os.path.join(output_folder, output_filename)
        cv2.imwrite(output_path, stacked_images)
        #print(f"Processed images saved to {output_path}")
        os.startfile(output_path)

    def stacked_images_for_markings(self,stacked_images, output_folder):
        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_filename = "stacked_images_" + current_time + ".jpg"
        output_path = os.path.join(output_folder, output_filename)
        cv2.imwrite(output_path, stacked_images)
        #print(f"Processed images saved to {output_path}")

    # Delete all files from the specific folder after saving the stacked image
    def delete_files_in_folder(self, folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}: {e}")
    
    def parse_answers_input(self,input_str):
        try:
            ans_list = eval(input_str)
            if not isinstance(ans_list, list) or not all(isinstance(ans, list) for ans in ans_list):
                raise ValueError("Invalid input format.")
            return ans_list
        except Exception:
            return None
        
    def process_omr(self):
        # Get answers list from input
        answers_input = self.formatted_combined_list
        answers_list = self.parse_answers_input(answers_input)
        info_message = "Invalid answers input. Please enter the answers in the specified format.\n"
        if answers_list is None:
            messagebox.showerror("ERROR !!!", info_message)
            return
        
        # Validate output folder path
        output_folder = self.output_folder_entry.get()
        if not output_folder or not os.path.isdir(output_folder):
            messagebox.showerror("ERROR !!!", "Invalid output folder path. Please select a valid output folder.")
            return
    
        total_score = 0
        total_correct = 0
        total_incorrect = 0
        total_skipped = 0
        processed_images = []
        processed_marks = []
    
        image_folder = r"C:\Users\kbnpa\Desktop\OMR Processing\Horizontal_Split"
        image_files = os.listdir(image_folder)
    
        # Loop over image paths and answers
        for i, image_file in enumerate(image_files):
            # Load the image
            image_path = os.path.join(image_folder, image_file)
    
            img = cv2.imread(image_path)
            img = cv2.resize(img, (self.widthImg, self.heightImg))
    
            processedImg = self.preprocessImage(img)
    
            imgContours = img.copy()
            imgFinal1 = img.copy()
            imgBiggestContours = img.copy()
            imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            imgBlur = cv2.GaussianBlur(imgGray, (5, 5), 1)
            imgCanny = cv2.Canny(imgBlur, 10, 50)
    
            contours, hierarchy = cv2.findContours(processedImg, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            cv2.drawContours(imgContours, contours, -1, (0, 255, 0), 10)
    
            rectCon = self.rectContour(contours)
            if not rectCon:
                # No valid contours found, skip processing this image
                print(f"Skipping image: {image_file} - No valid contour found.")
                continue
            
            biggestContour = self.getCornerPoints(rectCon[0])
    
            if biggestContour.size != 0:
                cv2.drawContours(imgBiggestContours, biggestContour, -1, (0, 0, 255), 20)
    
                biggestContour = self.reorder(biggestContour)
                pt1 = np.float32(biggestContour)
                pt2 = np.float32([[0, 0], [self.widthImg, 0], [0, self.heightImg], [self.widthImg, self.heightImg]])
                matrix1 = cv2.getPerspectiveTransform(pt1, pt2)
                imgwarpColored = cv2.warpPerspective(img, matrix1, (self.widthImg, self.heightImg))
                imgwarpGray = cv2.cvtColor(imgwarpColored, cv2.COLOR_BGR2GRAY)
                imgThresh = cv2.threshold(imgwarpGray, self.threshold_value, 255, cv2.THRESH_BINARY_INV)[1]
    
                boxes = self.splitBoxes(imgThresh)
    
                myPixelval = np.zeros((self.questions, self.choices))
                countC = 0
                countR = 0
    
                for image in boxes:
                    totalPixels = cv2.countNonZero(image)
                    myPixelval[countR][countC] = totalPixels
                    countC += 1
                    if countC == self.choices:
                        countR += 1
                        countC = 0
    
                myIndex = []
                for x in range (0,self.questions):
                    arr = myPixelval[x]
                    #print("Arr",arr)
                    myIndexVal = np.where(arr == np.amax(arr))
                    #print(myIndexVal[0])
                    myIndex.append(myIndexVal[0][0])
                #print(myIndex)
    
                answers = answers_list[i]
                if len(answers) != 5:
                    messagebox.showerror("ERROR", f"Invalid number of answers for OMR Sheet: {image_file}. Expected 50 answers for each sibject.")
                    continue
                ans = self.process_answers(answers)
    
    ######################################################################
                myResult = []                                           ##
                for x in range (0,self.questions):                      ##               
                    if myIndex[x] == 0:                                 ##               
                        myResult.append(1)                              ##               
                    if myIndex[x] == 1:                                 ##
                        myResult.append(2)                              ##
                    if myIndex[x] == 2:                                 ##
                        myResult.append(3)                              ##
                    if myIndex[x] == 3:                                 ##
                        myResult.append(4)                              ##
                    if myIndex[x] == 4:                                 ##
                        myResult.append(5)                              ##
                #print(myResult)                                        ##
    ######################################################################
    #######   MARKING SYSTEM                                            ## 
    ######################################################################
                grading = []                                            ##
                for x in range (0,self.questions):                      ##
                    if ans[x] == myIndex[x]:                            ##
                        grading.append(4)                               ##
                    else: grading.append(-1)                            ##
                #print(grading)                                         ##
                                                                        ##
                # NOTE: For Questions NOT Attempted                     ##
                countFifthChoice = 0                                    ##
                countFifthChoice = myResult.count(5)                    ##
                total_skipped +=countFifthChoice                        ##
                #print("No. of questions skipped",countFifthChoice)     ##
                markedFifthChoice = []                                  ##
                for i in range(len(myResult)):                          ##
                    if myResult[i] == 5:                                ##
                        markedFifthChoice.append(i+1)                   ##
                # NOTE: To know which questions are skipped:-           ##
                #print("Question skipped", markedFifthChoice)           ##
                                                                        ##
                # NOTE: For Incorrect Questions                         ##
                countIncorrect = 0                                      ##
                countIncorrect = grading.count(-1) - countFifthChoice   ##
                total_incorrect += countIncorrect                       ##
                #print("No. of incorrect questions", countIncorrect)    ##
                markedIncorrect = []                                    ##
                for i in range(len(grading)):                           ##
                    if grading[i] == -1 and myResult[i] != 5:           ##
                        markedIncorrect.append(i+1)                     ##
                # NOTE: To know which questions are wrong:-             ##
                #print("Incorrect questions", markedIncorrect)          ##
                                                                        ##
                # NOTE: For Correct Questions                           ##
                countCorrect = 0                                        ##
                countCorrect = grading.count(4)                         ##
                total_correct += countCorrect                           ##
                #print("No. of Correct questions", countCorrect)        ##
                markedCorrect = []                                      ##
                for i in range(len(grading)):                           ##
                    if grading[i] == 4:                                 ##
                        markedCorrect.append(i+1)                       ##
                # NOTE: To know which questions are correct:-           ##
                #print("Correct questions", markedCorrect)              ##
                                                                        ##
                # NOTE: Score Calculation                               ##
                score = sum(grading) + countFifthChoice                 ##
                total_score += score                                    ##
                #print(score)                                           ##
                #print("Total Marks Obtained:", total_score)            ##
    ######################################################################
    
                imgResult = imgwarpColored.copy()
                imgResult = self.showAnswers(imgResult, myIndex, grading, ans, self.questions, self.choices)
                imgRawDrawing = np.zeros_like(imgwarpColored)
                imgRawDrawing = self.showAnswers(imgRawDrawing, myIndex, grading, ans, self.questions, self.choices)
                invmatrix1 = cv2.getPerspectiveTransform(pt2, pt1)
                imgInvWarp = cv2.warpPerspective(imgRawDrawing, invmatrix1, (self.widthImg, self.heightImg))
    
                imgFinal1 = cv2.addWeighted(imgFinal1, 1, imgInvWarp, 1, 0)
    
            imgBlank = np.zeros_like(img)
            imageArry = [
                [imgGray, imgBlur, imgCanny, imgContours],
                [imgBiggestContours, imgwarpColored, imgThresh, imgResult],
                [imgRawDrawing, imgInvWarp, imgFinal1, imgBlank]
            ]
            
            processed_images.append(imgFinal1)
            processed_marks.append(imgInvWarp)
    
            answers = answers_list[i]
            ans = self.process_answers(answers)
        
            # Display the result
            imgStacked = self.stackImages(imageArry, 0.2)
            #cv2.imshow("Check",imgStacked)
            #cv2.imshow("check", imgThresh) 
    
        # Display the result on the GUI
        label_text = f"Total questions attended: {total_correct + total_incorrect}\n" \
                     f"Total questions skipped: {total_skipped}\nTotal correct questions: {total_correct}\n" \
                     f"Total incorrect questions: {total_incorrect}\nTotal marks attended: {4 * total_correct + 4 * total_incorrect}\n" \
                     f"Total positive marks: {4 * total_correct}\nTotal negative marks: {total_incorrect}\n" \
                     f"Total Marks Obtained: {total_score}"
        
        label_text_image = f"Total questions attended: {total_correct + total_incorrect}\n" \
                           f"Total questions skipped: {total_skipped}\nTotal correct questions: {total_correct}\n" \
                           f"Total incorrect questions: {total_incorrect}\nTotal marks attended: {4 * total_correct + 4 * total_incorrect}\n" \
                           f"Total positive marks: {4 * total_correct}\nTotal negative marks: {total_incorrect}\n" \
                           f"Total Marks Obtained: {total_score}\n\n" \
                           f"\n\n" \
                           f"Incorrect attempts: RED Dots.  \n" \
                           f"Correct attempts: GREEN Dots.  \n" \
                           f"Skipped questions: YELLOW Dots.\n"
        
        #stacked_images = self.stack_images_col_wise(processed_images, self.subjects)
        stacked_images = self.stack_images_col_wise(processed_images, self.subjects, label_text_image)
        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_filename = "stacked_images_" + current_time + ".jpg"

#        stacked_images1 = self.stack_images_col_wise_1(processed_images, self.subjects)####################################
#        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
#        output_filename1 = "stacked_images_" + current_time + ".jpg"
    
        self.canvas2.delete("all")
        self.start_stopwatch()
        self.animate_progress(100)
        self.draw_stopwatch(4500)
        self.update_stopwatch()
        self.start_stopwatch()
        self.stop_stopwatch()
        self.canvas2.delete("all")

        # Update the text in the text box
        self.text_box.config(state=tk.NORMAL)  # Allow modification of the text
        self.text_box.delete("1.0", tk.END)    # Clear the existing content
        self.text_box.insert(tk.END, label_text)  # Insert the new content
        self.text_box.config(state=tk.DISABLED)  # Disable modification of the text

        # Save the stacked images (without spacing)
        output_folder = self.output_folder_entry.get()
        self.save_stacked_images(stacked_images, output_folder)

#        # Save the stacked images (without spacing)
#        output_folder = self.output_folder_entry.get()######################################################
#        self.save_stacked_images(stacked_images1, output_folder)

        # Call the function to delete files from the specified folder
        folder_to_delete_files = r"C:\Users\kbnpa\Desktop\OMR Processing\Processed_OMR"
        self.delete_files_in_folder(folder_to_delete_files)

        output_folder2 = r"C:\Users\kbnpa\Desktop\OMR Processing\Processed_OMR"
        self.stacked_images_for_markings(stacked_images, output_folder2)

        cv2.waitKey(0)
        cv2.destroyAllWindows()

# Create the GUI instance
app = OMR_Scanner()
app.mainloop()