import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk
import tkinter.font as tkFont # Import the font module
import sys
import os
import threading
import time
import subprocess
import re

print("--- Script has started execution! ---")
print("--- Standard library imports complete ---")

# --- Selenium imports ---
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

print("--- Selenium imports complete ---")

# --- Helper function for web automation ---
def automate_web_actions(url, webdriver_path, command_to_type, log_func, set_static_ip_func,
                         target_wifi_ssid_web, target_wifi_password_web, resume_event, app_instance):
    """
    Automates web actions. Output is sent via log_func.
    set_static_ip_func is a callback to set the static IP.
    target_wifi_ssid_web and target_wifi_password_web are for the device's web UI.
    resume_event: A threading.Event to signal when to resume script.
    app_instance: Reference to the Application instance to update GUI elements.
    """
    service = Service(webdriver_path)
    driver = None
    try:
        log_func("\n--- Starting Web Automation ---")
        driver = webdriver.Edge(service=service)
        driver.maximize_window()

        log_func(f"Starting script by opening website: {url}")
        driver.get(url)

        # --- Step 1: Click the "Files" link ---
        log_func("Waiting for the 'Files' link to appear...")
        files_link = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Files')]"))
        )
        log_func("'Files' link found. Clicking it...")
        files_link.click()
        log_func("Clicked the 'Files' link.")
        time.sleep(2)

        # --- Step 2: Loop to delete all "ECache" files ---
        log_func("\n--- Deleting ECache files ---")
        delete_count = 0
        max_attempts_overall = 100

        try:
            driver.save_screenshot("before_ecache_deletion.png")
            log_func("Screenshot 'before_ecache_deletion.png' saved.")
        except Exception as e:
            log_func(f"Could not save screenshot before ECache deletion: {e}")

        for attempt in range(1, max_attempts_overall + 1):
            log_func(f"\n--- ECache Deletion Attempt {attempt} ---")
            time.sleep(0.5)

            try:
                ecache_file_elements = driver.find_elements(By.XPATH,
                    "//div[contains(@class, 'file') and contains(@class, 'fs-file') and contains(@class, 'deletable') and .//a[starts-with(@href, 'http://setup.com/ECache_') or starts-with(text(), 'ECache_')]]")

                num_ecache_files_found = len(ecache_file_elements)
                log_func(f"Found {num_ecache_files_found} ECache file(s) on page.")

                if not ecache_file_elements:
                    log_func("No more ECache files found. Exiting deletion loop.")
                    break

                log_func(f"Found {len(ecache_file_elements)} ECache file(s) to delete in this pass.")
                target_ecache_file_element = ecache_file_elements[0]
                parent_data_id = target_ecache_file_element.get_attribute("data-id")
                log_func(f"Targeting ECache file with data-id='{parent_data_id}'.")

                delete_button_locator = (By.XPATH, f".//div[@class='status' and @data-id='{parent_data_id}']")

                log_func(f"Attempting to click delete for ECache file data-id='{parent_data_id}'...")
                individual_delete_button = WebDriverWait(target_ecache_file_element, 5).until(
                    EC.element_to_be_clickable(delete_button_locator)
                )
                individual_delete_button.click()
                log_func("Successfully clicked individual ECache delete button.")
                time.sleep(1)

                log_func("Waiting for delete confirmation modal to appear...")
                confirm_delete_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'modal-primary') and contains(text(), 'Delete')]"))
                )
                log_func("Delete confirmation button found. Attempting to click to confirm...")
                confirm_delete_button.click()
                log_func("Successfully clicked delete confirmation button.")
                delete_count += 1
                time.sleep(3)

            except StaleElementReferenceException:
                log_func("StaleElementReferenceException caught. Element reference is no longer valid, likely due to DOM change. Retrying this deletion attempt.")
                try:
                    driver.save_screenshot(f"stale_element_error_attempt_{attempt}.png")
                except Exception as e:
                    log_func(f"Could not save screenshot for stale element error: {e}")
                time.sleep(1)
                continue
            except TimeoutException as te:
                log_func(f"TimeoutException caught during ECache deletion: {te}")
                log_func("Could not find element within specified time. This might mean all ECache files are already deleted, or the locator is wrong, or elements are not becoming clickable.")
                try:
                    driver.save_screenshot(f"timeout_error_attempt_{attempt}.png")
                except Exception as e:
                    log_func(f"Could not save screenshot for timeout error: {e}")
                break
            except Exception as e:
                log_func(f"An unexpected error occurred during ECache deletion (Attempt {attempt}): {e}")
                import traceback
                log_func(traceback.format_exc())
                try:
                    driver.save_screenshot(f"general_error_attempt_{attempt}.png")
                except Exception as e_ss:
                    log_func(f"Could not save screenshot for general error: {e_ss}")
                time.sleep(2)

        log_func(f"\nFinished ECache deletion loop. Total deleted: {delete_count}")
        if attempt >= max_attempts_overall and num_ecache_files_found > 0:
            log_func(f"Warning: Reached maximum ECache deletion attempts ({max_attempts_overall}). Some ECache files might remain.")

        time.sleep(2)

        # --- Step 3: Click the "Connect" link ---
        log_func("\nWaiting for the 'Connect' link to appear...")
        connect_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Connect"))
        )
        log_func("'Connect' link found. Clicking it...")
        connect_link.click()
        log_func("Clicked the 'Connect' link.")

        # --- NEW: Automate Wi-Fi connection on the device's web interface ---
        log_func("\n--- Automating Wi-Fi connection on the device's web interface ---")
        try:
            # Wait for the network list to appear (the 'networks' div)
            log_func("Waiting for the Wi-Fi network list to appear...")
            networks_container = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, "networks"))
            )
            log_func("Wi-Fi network list found.")

            # Find the specific network by its SSID text
            log_func(f"Searching for target Wi-Fi network: '{target_wifi_ssid_web}'...")
            target_network_element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, f"//div[@class='network']/div[@class='ssid'][text()='{target_wifi_ssid_web}']/ancestor::div[@class='network']"))
            )
            log_func(f"Found target network '{target_wifi_ssid_web}'. Clicking it...")
            target_network_element.click()
            log_func(f"Clicked on network '{target_wifi_ssid_web}'.")
            time.sleep(2) # Give time for the password input to appear

            # Now, the password input field should appear
            log_func("Waiting for password input field to appear...")
            password_input_field = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='password'][type='password']"))
            )
            password_input_field.clear()
            password_input_field.send_keys(target_wifi_password_web)
            log_func("Typed password into the field.")

            # Click the "Connect" button that saves the credentials (class="btn btn-lg save")
            log_func("Attempting to click the 'Connect' button (save button)...")
            web_connect_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-lg.save"))
            )
            web_connect_button.click()
            log_func("Clicked the 'Connect' button to save credentials.")

            log_func("Giving device time to process connection and reboot/reconnect (60-90 seconds recommended for network change and potential reboot).")
            time.sleep(90) # Increased sleep time for network changes and potential device reboot

        except TimeoutException as te:
            log_func(f"TimeoutException caught during web Wi-Fi connection: {te}")
            log_func("Could not find Wi-Fi network, password input, or connect button on web page within time. The page structure might have changed or network not found.")
            log_func(f"Ensure target Wi-Fi '{target_wifi_ssid_web}' is visible on the scan page.")
        except NoSuchElementException as nse:
            log_func(f"NoSuchElementException caught during web Wi-Fi connection: {nse}")
            log_func("A required element was not found. Check the provided HTML and locators.")
        except Exception as e:
            log_func(f"An unexpected error occurred during web Wi-Fi connection: {e}")
            import traceback
            log_func(traceback.format_exc())

        # --- PAUSE POINT: Wait for user to manually reconnect PC to JuiceNet ---
        log_func("\n--- DEVICE CONFIGURATION COMPLETE. ---")
        log_func("Please MANUALLY RECONNECT your PC to the JuiceNet network (e.g., JuiceNet-BC9) through your system's Wi-Fi settings.")
        log_func("Once reconnected, click the 'Resume Script (Connected to JuiceNet)' button in the GUI.")

        # Enable the resume button and disable others on the GUI thread
        app_instance.after(0, app_instance.pages["Page3_Automation"].enable_resume_button)
        app_instance.after(0, app_instance.get_and_display_current_ip_threaded_wrapper) # Refresh PC IP display

        resume_event.wait() # This will block the automation thread until the event is set
        log_func("Resume signal received. Script continuing...")

        # Disable resume button and re-enable others after resuming
        app_instance.after(0, app_instance.pages["Page3_Automation"].disable_resume_button)

        # --- Set static IP after reconnection and before refreshing browser ---
        # This call is intentionally here, as it's part of the automation flow.
        # The set_static_ip_func now has internal checks to only apply to JuiceNet.
        log_func("Attempting to set PC static IP now (conditional on JuiceNet connection)...")
        set_static_ip_func() # Call the function to set static IP
        time.sleep(5) # Give a moment for IP change to propagate

        log_func("Refreshing browser to ensure connection...")
        driver.refresh()
        time.sleep(5)

        # --- Step 4: Click "Console" link ---
        log_func("\nWaiting for the 'Console' link to appear...")
        console_link = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Console"))
        )
        log_func("'Console' link found. Clicking it...")
        console_link.click()
        log_func("Clicked the 'Console' link.")
        time.sleep(2)

        # --- Step 5: Type into the console input ---
        log_func("Waiting for the console input field to appear in the modal...")
        console_input_field = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "cmdline"))
        )
        log_func(f"Console input field found. Typing command: '{command_to_type}'")
        console_input_field.send_keys(command_to_type)
        console_input_field.send_keys(webdriver.Keys.ENTER)
        log_func("Typed command and pressed ENTER.")
        time.sleep(5)

    except Exception as e:
        log_func(f"\n--- An unhandled error occurred during web automation: {e} ---")
        import traceback
        log_func(traceback.format_exc())
        app_instance.after(0, lambda: messagebox.showerror("Error", f"An error occurred: {e}\nCheck the log within the GUI for details."))
    finally:
        if driver:
            log_func("Closing the browser.")
            driver.quit()
        log_func("\n--- Web Automation Process Finished. ---")
        app_instance.after(0, app_instance.automation_finished_callback)


# --- Page Frame Definitions ---

class BasePage(ttk.Frame): # Use ttk.Frame
    """Base class for all pages to provide common functionality."""
    def __init__(self, parent, controller):
        ttk.Frame.__init__(self, parent) # Use ttk.Frame __init__
        self.controller = controller
        self.columnconfigure(0, weight=1) # Allow content to expand
        self.rowconfigure(0, weight=1) # Allow content to expand
        self.next_button = None # Initialize to None
        self.prev_button = None # Initialize to None

    def setup_navigation_buttons(self, next_page_class=None, prev_page_class=None, next_button_text="Next"):
        """Adds navigation buttons to the bottom of the page."""
        nav_frame = ttk.Frame(self) # Use ttk.Frame
        nav_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        nav_frame.columnconfigure(0, weight=1)
        nav_frame.columnconfigure(1, weight=1)

        if prev_page_class:
            self.prev_button = ttk.Button(nav_frame, text="Back", command=lambda: self.controller.show_frame(prev_page_class), style='TButton') # Apply default button style
            self.prev_button.pack(side=tk.LEFT, padx=5)

        if next_page_class:
            self.next_button = ttk.Button(nav_frame, text=next_button_text, command=lambda: self.controller.show_frame(next_page_class), style='Accent.TButton') # Apply accent style
            self.next_button.pack(side=tk.RIGHT, padx=5)


class Page1_PCCheck(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_widgets()
        self.controller.log_message("Page 1: Connect PC to JuiceNet network.")

    def create_widgets(self):
        page_frame = ttk.LabelFrame(self, text="Step 1: Connect PC to Juice Box Wi-Fi", style='Page.TLabelframe') # Apply style
        page_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # PC's Wi-Fi List and Connect
        ttk.Label(page_frame, text="Available PC Networks:", style='Page.TLabel').pack(pady=(0, 5)) # Apply style
        self.pc_wifi_listbox = tk.Listbox(page_frame, height=5, width=40,
                                          bg='#2C2F33', fg='white', selectbackground='#7289DA', selectforeground='white',
                                          highlightbackground='#505050', highlightcolor='#7289DA', bd=0, relief='flat',
                                          font=self.controller.default_font) # Manual styling for tk.Listbox, added font
        self.pc_wifi_listbox.pack(pady=(0, 5))

        self.pc_wifi_refresh_button = ttk.Button(page_frame, text="Refresh PC Wi-Fi List",
                                                command=self.controller.populate_pc_wifi_list_threaded_wrapper, style='TButton') # Apply style
        self.pc_wifi_refresh_button.pack(pady=5)

        ttk.Label(page_frame, text="PC Wi-Fi Password (if required):", style='Page.TLabel').pack(pady=(5, 0)) # Apply style
        self.pc_wifi_password_entry = ttk.Entry(page_frame, show="*", width=40, style='TEntry') # Apply style
        self.pc_wifi_password_entry.pack(pady=(0, 5))

        self.connect_pc_wifi_button = ttk.Button(page_frame, text="Connect PC to Selected Wi-Fi",
                                                 command=self._on_connect_pc_wifi_button_click, state=tk.DISABLED, style='Accent.TButton') # Apply style
        self.connect_pc_wifi_button.pack(pady=5)

        self.setup_navigation_buttons(next_page_class=Page2_DeviceSetup)

        # Bind listbox selection to enable connect button
        self.pc_wifi_listbox.bind("<<ListboxSelect>>", self._on_listbox_select)

    def _on_listbox_select(self, event=None):
        if self.pc_wifi_listbox.curselection():
            self.connect_pc_wifi_button.config(state=tk.NORMAL)
        else:
            self.connect_pc_wifi_button.config(state=tk.DISABLED)

    def _on_connect_pc_wifi_button_click(self):
        selected_index = self.pc_wifi_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("No Selection", "Please select a Wi-Fi network from the PC's list.")
            return

        ssid_line = self.pc_wifi_listbox.get(selected_index[0])
        ssid_match = re.match(r"^\d+\.\s*(.*)", ssid_line)
        if ssid_match:
            selected_ssid = ssid_match.group(1).strip()
        else:
            selected_ssid = ssid_line.strip()

        password = self.pc_wifi_password_entry.get()
        self.controller.connect_pc_to_wifi(selected_ssid, password)

    def on_show(self):
        """Called when this page is displayed."""
        self.controller.populate_pc_wifi_list_threaded_wrapper()
        self.controller.get_and_display_current_ip_threaded_wrapper()


class Page2_DeviceSetup(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_widgets()
        self.controller.log_message("Page 2: Set PC Static IP and configure Device's Wi-Fi.")

    def create_widgets(self):
        page_frame = ttk.LabelFrame(self, text="Step 2: PC IP & Device Wi-Fi Setup", style='Page.TLabelframe') # Apply style
        page_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # PC's IP Configuration Buttons
        ttk.Button(page_frame, text="Set PC Static IP (10.10.10.2) [JuiceNet only]",
                  command=self.controller.set_static_ip_threaded_wrapper, style='TButton').pack(pady=10) # Apply style

        # Input fields for Web Wi-Fi connection (for the DEVICE)
        ttk.Label(page_frame, text="Target Wi-Fi Name (for DEVICE to connect to):", style='Page.TLabel').pack(pady=(10, 0)) # Apply style
        self.target_device_wifi_ssid_entry = ttk.Entry(page_frame, width=30, style='TEntry') # Apply style
        self.target_device_wifi_ssid_entry.pack(pady=(0, 5))
        self.target_device_wifi_ssid_entry.insert(0, "YourHomeNetwork") # Suggest a default

        ttk.Label(page_frame, text="Target Wi-Fi Password (for DEVICE):", style='Page.TLabel').pack(pady=(5, 0)) # Apply style
        self.target_device_wifi_password_entry = ttk.Entry(page_frame, show="*", width=30, style='TEntry') # Apply style
        self.target_device_wifi_password_entry.pack(pady=(0, 10))

        self.setup_navigation_buttons(next_page_class=Page3_Automation, prev_page_class=Page1_PCCheck)

    def on_show(self):
        """Called when this page is displayed."""
        self.controller.get_and_display_current_ip_threaded_wrapper()


class Page3_Automation(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_widgets()
        self.controller.log_message("Page 3: Start Web Automation.")

    def create_widgets(self):
        page_frame = ttk.LabelFrame(self, text="Step 3: Start Device Automation", style='Page.TLabelframe') # Apply style
        page_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Start Automation Button
        self.start_button = ttk.Button(page_frame, text="Start Automation", command=self._start_automation_wrapper, width=30, style='Accent.TButton') # Apply style
        self.start_button.pack(pady=20)

        # New Resume Button
        self.resume_script_button = ttk.Button(page_frame, text="Resume Script (Connected to JuiceNet)",
                                              command=self.controller.resume_automation, width=30, state=tk.DISABLED,
                                              style='Resume.TButton') # Custom style for resume button
        self.resume_script_button.pack(pady=10)

        self.connect_wifi_button = ttk.Button(
            page_frame,
            text="Connect PC to Wi-Fi",
            command=self._connect_wifi_step3,
            width=30,
            style='Accent.TButton'
        )
        self.connect_wifi_button.pack(pady=10)

        # Use the standard navigation setup, pointing "Next" to Page 4
        self.setup_navigation_buttons(prev_page_class=Page2_DeviceSetup, next_page_class=Page4_Cleanup)
        # Initially disable the 'Next' button until automation is complete
        if self.next_button:
            self.next_button.config(state=tk.DISABLED)

    def _start_automation_wrapper(self):
        # Get values from Page 2's entries
        page2 = self.controller.frames["Page2_DeviceSetup"]
        target_device_wifi_ssid = page2.target_device_wifi_ssid_entry.get().strip()
        target_device_wifi_password = page2.target_device_wifi_password_entry.get().strip()

        if not target_device_wifi_ssid:
            messagebox.showwarning("Missing Input", "Please enter the Target Wi-Fi Name for the DEVICE on Page 2.")
            self.controller.show_frame(Page2_DeviceSetup) # Go back to Page 2
            return

        self.controller.start_automation(target_device_wifi_ssid, target_device_wifi_password)

    def _connect_wifi_step3(self):
        page2 = self.controller.frames["Page2_DeviceSetup"]
        ssid = page2.target_device_wifi_ssid_entry.get().strip()
        password = page2.target_device_wifi_password_entry.get().strip()
        self.controller.connect_pc_to_wifi(ssid, password)

    def enable_resume_button(self):
        """Enables the resume button and disables start button."""
        self.resume_script_button.config(state=tk.NORMAL)
        self.start_button.config(state=tk.DISABLED)
        if self.next_button:
            self.next_button.config(state=tk.DISABLED) # Disable next until automation is done

    def disable_resume_button(self):
        """Disables the resume button and re-enables start button."""
        self.resume_script_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)
        # Note: next_button state will be managed by automation_finished_callback

    def on_show(self):
        """Called when this page is displayed."""
        self.controller.get_and_display_current_ip_threaded_wrapper()
        # Reset button states if returning to this page
        if not (self.controller.automation_thread and self.controller.automation_thread.is_alive()):
            self.start_button.config(state=tk.NORMAL)
            self.resume_script_button.config(state=tk.DISABLED)
            # Re-enable next button if automation is already finished
            if self.next_button:
                if self.controller.automation_finished_flag:
                    self.next_button.config(state=tk.NORMAL)
                else:
                    self.next_button.config(state=tk.DISABLED)
        else: # Automation is running
            self.start_button.config(state=tk.DISABLED)
            # Resume button state depends on if it's currently paused and waiting
            if self.controller.resume_automation_event.is_set(): # Automation has resumed/finished this step
                 self.resume_script_button.config(state=tk.DISABLED)
            else: # Automation is waiting for resume
                 self.resume_script_button.config(state=tk.NORMAL)
            if self.next_button:
                self.next_button.config(state=tk.DISABLED)


class Page4_Cleanup(BasePage):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.create_widgets()
        self.controller.log_message("Page 4: Revert PC IP settings.")

    def create_widgets(self):
        page_frame = ttk.LabelFrame(self, text="Step 4: Cleanup - Revert PC IP", style='Page.TLabelframe') # Apply style
        page_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        ttk.Label(page_frame, text="It is highly recommended to revert your PC's IP settings back to DHCP after finishing.",
                 wraplength=400, justify=tk.CENTER, style='Page.TLabel').pack(pady=20) # Apply style

        ttk.Button(page_frame, text="Revert PC IP to DHCP",
                  command=self.controller.revert_ip_to_dhcp_threaded_wrapper, width=30,
                  style='Warning.TButton').pack(pady=10) # Custom style for warning button

        self.setup_navigation_buttons(prev_page_class=Page3_Automation)

        # Override the "Next" button from BasePage to be the Exit button
        # Access the navigation frame where the button would be packed
        nav_frame = None
        for widget in self.winfo_children():
            if isinstance(widget, ttk.Frame) and widget.winfo_children():
                # Check if it's the navigation frame by looking for buttons
                if any(isinstance(child, ttk.Button) for child in widget.winfo_children()):
                    nav_frame = widget
                    break

        if self.next_button and nav_frame:
            self.next_button.destroy() # Remove default next button
        
        # Create a new Exit button in the same nav_frame if found, otherwise pack it directly
        if nav_frame:
            self.exit_button = ttk.Button(nav_frame, text="Exit Application", command=self.controller.on_exit, width=20, style='Accent.TButton') # Apply style
            self.exit_button.pack(side=tk.RIGHT, padx=5)
        else:
            # Fallback if nav_frame not found (shouldn't happen with current setup)
            self.exit_button = ttk.Button(self, text="Exit Application", command=self.controller.on_exit, width=20, style='Accent.TButton')
            self.exit_button.pack(side=tk.RIGHT, padx=5, anchor=tk.SE) # Pack at bottom right


    def on_show(self):
        """Called when this page is displayed."""
        self.controller.get_and_display_current_ip_threaded_wrapper()


# --- Main GUI Application ---
class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Web Automation Tool - Step-by-Step")
        self.geometry("800x700") # Adjust geometry for more space

        # --- Styling Configuration ---
        self.style = ttk.Style(self)
        self.style.theme_use("clam") # A good base theme to customize for modern look

        # Define colors inspired by VoltiE
        self.bg_color = '#1A1C20'       # Very dark blue-grey, almost black
        self.fg_color = 'white'         # White text
        self.accent_color = '#7289DA'   # Light purple/blue (from original, fits well)
        self.button_hover_color = '#5A6EA8' # Darker accent for hover
        self.warning_color = '#DC3545'  # Red for warnings/revert
        self.resume_color = '#28A745'   # Green for resume

        # Configure root window background
        self.configure(bg=self.bg_color)

        # Global Font - Try "Segoe UI" first, then common sans-serif
        # Corrected: Use tkFont.families()
        self.default_font_family = "Segoe UI" if "Segoe UI" in tkFont.families() else "Arial"
        self.default_font = (self.default_font_family, 10)
        self.bold_font = (self.default_font_family, 10, "bold")
        self.log_font = ("Consolas", 9) # Monospace for logs

        # Configure global styles for all ttk widgets
        self.style.configure('.', font=self.default_font, background=self.bg_color, foreground=self.fg_color)

        # Style for TFrame (general frames and container)
        self.style.configure('TFrame', background=self.bg_color)

        # Style for TLabelframe (section headers)
        self.style.configure('TLabelframe', background=self.bg_color, foreground=self.accent_color,
                             font=self.bold_font, borderwidth=1, relief='solid', padding=[10, 5, 10, 10])
        # Inner label of the Labelframe (the text itself)
        self.style.configure('TLabelframe.Label', background=self.bg_color, foreground=self.accent_color,
                             font=self.bold_font, padding=(5, 2)) # Adjust padding for label inside frame

        # Style for TLabel (general text labels)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, padding=5)
        self.style.configure('Page.TLabel', background=self.bg_color, foreground=self.fg_color, padding=5) # Specific label style for pages

        # Style for TEntry (input fields)
        self.style.configure('TEntry', fieldbackground='#3C3F41', foreground='white', insertbackground='white',
                             borderwidth=1, relief='flat', padding=5)
        self.style.map('TEntry', fieldbackground=[('focus', '#4C4F51')])

        # Base TButton style (for "Back", "Refresh PC Wi-Fi List", "Set PC Static IP")
        self.style.configure('TButton', background='#3C3F41', foreground=self.fg_color, borderwidth=0, relief='flat', padding=8) # Increased padding
        self.style.map('TButton',
                       background=[('active', self.button_hover_color), ('disabled', '#2A2C2F')],
                       foreground=[('active', 'white'), ('disabled', '#808080')],
                       relief=[('pressed', 'sunken'), ('!pressed', 'flat')]) # Ensure flat look even after press

        # Accent Button Style (for "Connect PC", "Start Automation", "Next", "Exit Application")
        self.style.configure('Accent.TButton', background=self.accent_color, foreground='white', borderwidth=0, relief='flat', padding=10) # More padding for primary
        self.style.map('Accent.TButton',
                       background=[('active', self.button_hover_color), ('disabled', '#404040')],
                       foreground=[('active', 'white'), ('disabled', '#A0A0A0')],
                       relief=[('pressed', 'sunken'), ('!pressed', 'flat')])

        # Warning Button Style (for "Revert PC IP to DHCP")
        self.style.configure('Warning.TButton', background=self.warning_color, foreground='white', borderwidth=0, relief='flat', padding=10)
        self.style.map('Warning.TButton',
                       background=[('active', '#B82030'), ('disabled', '#404040')],
                       foreground=[('active', 'white'), ('disabled', '#A0A0A0')],
                       relief=[('pressed', 'sunken'), ('!pressed', 'flat')])

        # Resume Button Style (green for success/go)
        self.style.configure('Resume.TButton', background=self.resume_color, foreground='white', borderwidth=0, relief='flat', padding=10)
        self.style.map('Resume.TButton',
                       background=[('active', '#208535'), ('disabled', '#404040')],
                       foreground=[('active', 'white'), ('disabled', '#A0A0A0')],
                       relief=[('pressed', 'sunken'), ('!pressed', 'flat')])


        # Initialize core attributes *before* creating widgets that might use them
        self.automation_thread = None
        self.TARGET_URL = "http://setup.com"
        self.COMMAND_TO_EXECUTE = "dfuu -i wlan --multi"

        # --- Wi-Fi Adapter Name for IP checks/settings ---
        self.WIFI_ADAPTER_NAME = "Wi-Fi"
        # Pattern to identify JuiceNet device's Wi-Fi.
        self.JUICENET_SSID_PATTERN = "JuiceNet"

        # --- Static IP Configuration ---
        self.TARGET_STATIC_IP = "10.10.10.2"
        self.TARGET_SUBNET_MASK = "255.255.255.0"
        self.TARGET_GATEWAY = "10.10.10.1"
        self.TARGET_DNS = "10.10.10.1"

        # Flag to track if this program instance set a static IP
        self.ip_was_set_statically = False
        self.automation_finished_flag = False # New flag to indicate automation completion

        # Threading Event for pausing/resuming automation
        self.resume_automation_event = threading.Event()

        # Determine driver path
        if getattr(sys, 'frozen', False):
            bundle_dir = sys._MEIPASS
            self.EDGE_DRIVER_PATH = os.path.join(bundle_dir, "msedgedriver.exe")
        else:
            # IMPORTANT: Change this to your actual msedgedriver.exe path for development
            self.EDGE_DRIVER_PATH = "C:/Users/ruano/Desktop/msedgedriver.exe"
            # In a real deployed app, you'd want a more robust way to find/distribute the driver

        self.frames = {}
        self.current_frame = None

        self.create_global_widgets() # Widgets that persist across pages
        self.create_pages() # Create the page frames

        # Set up a protocol for handling window close (X button)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Show the first page
        self.show_frame(Page1_PCCheck)

        self.log_message(f"Running as {'bundled executable' if getattr(sys, 'frozen', False) else 'Python script'}. Driver path: {self.EDGE_DRIVER_PATH}")
        if not os.path.exists(self.EDGE_DRIVER_PATH):
            messagebox.showerror("Driver Error", f"""MSEdgeDriver not found at {self.EDGE_DRIVER_PATH}
Please ensure 'msedgedriver.exe' is in the correct location or bundled correctly.""")
            self.log_message("--- Driver not found, attempting to quit ---")
            self.quit()
        self.log_message("--- Driver path check passed ---")
        self.log_message("\n--- REMEMBER TO RUN THIS SCRIPT AS ADMINISTRATOR FOR PC IP CHANGES! ---")

    def create_global_widgets(self):
        """Widgets that remain visible regardless of the current page."""

        # Top Frame for PC's Current IP Display
        top_info_frame = ttk.Frame(self, style='TFrame', relief=tk.GROOVE) # Use ttk.Frame
        top_info_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(top_info_frame, text="Current PC IP Address:", style='TLabel').grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.current_ip_label = ttk.Label(top_info_frame, text="N/A", font=self.bold_font, foreground=self.accent_color, style='TLabel') # Use accent color
        self.current_ip_label.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(top_info_frame, text="Connected SSID:", style='TLabel').grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.connected_ssid_label = ttk.Label(top_info_frame, text="N/A", font=self.bold_font, foreground='#8BC34A', style='TLabel') # Green for success
        self.connected_ssid_label.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # Placeholder for dynamic page frames
        self.container = ttk.Frame(self, style='TFrame') # Use ttk.Frame
        self.container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # ScrolledText widget for output (fills the remaining space at the bottom)
        self.text_area = scrolledtext.ScrolledText(self, wrap=tk.WORD, height=12, width=90,
                                                   bg='#1E1E1E', fg='#E0E0E0', insertbackground='white',
                                                   font=self.log_font, bd=0, relief='flat') # Dark background, light text for log
        self.text_area.pack(padx=10, pady=(0, 10), fill=tk.BOTH, expand=True)

    def create_pages(self):
        """Instantiates all page frames and stores them."""
        for F in (Page1_PCCheck, Page2_DeviceSetup, Page3_Automation, Page4_Cleanup):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

    def show_frame(self, page_class):
        """Brings a specific page frame to the front."""
        frame = self.frames[page_class.__name__]
        frame.tkraise()
        self.current_frame = frame
        if hasattr(frame, 'on_show'):
            frame.on_show() # Call a method on the page when it's shown for refreshing content

    def log_message(self, message):
        """Appends a message to the text_area widget."""
        try:
            self.text_area.insert(tk.END, message + "\n")
            self.text_area.see(tk.END) # Auto-scroll to the end
            self.text_area.update_idletasks() # Update GUI immediately
        except AttributeError:
            print(f"LOG_ERROR: text_area not ready: {message}")


    def _create_wifi_profile_xml(self, ssid, password):
        """Generates an XML string for a WPA2-PSK Wi-Fi profile."""
        xml_template = """<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{ssid}</name>
    <SSIDConfig>
        <SSID>
            <name>{ssid}</name>
        </SSID>
    </SSIDConfig>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authAndCiphers>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authAndCiphers>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{password}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>"""
        return xml_template.format(ssid=ssid, password=password)


    def populate_pc_wifi_list_threaded_wrapper(self):
        """Starts the PC's Wi-Fi scan in a separate thread to keep GUI responsive."""
        page1 = self.frames["Page1_PCCheck"]
        page1.pc_wifi_refresh_button.config(state=tk.DISABLED)
        page1.connect_pc_wifi_button.config(state=tk.DISABLED)
        self.log_message("\n--- Initiating PC Wi-Fi scan in background... ---")
        page1.pc_wifi_listbox.delete(0, tk.END)
        threading.Thread(target=self._scan_pc_wifi_networks_threaded, daemon=True).start()

    def _scan_pc_wifi_networks_threaded(self):
        """Performs the netsh scan for PC's Wi-Fi in a separate thread."""
        ssids = []
        error_message = None
        try:
            command = "chcp 65001 && netsh wlan show networks"
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                shell=True
            )
            output = result.stdout
            ssids = re.findall(r"SSID \d+ : (.*)", output, re.IGNORECASE)

        except subprocess.CalledProcessError as e:
            error_message = f"Error running netsh command for PC Wi-Fi scan: {e}\n  Stderr: {e.stderr.strip()}"
        except FileNotFoundError:
            error_message = "Error: 'netsh' command not found. This command is specific to Windows."
        except Exception as e:
            error_message = f"An unexpected error occurred during PC Wi-Fi scan: {e}"

        self.after(0, self._update_pc_wifi_list_gui, ssids, error_message)

    def _update_pc_wifi_list_gui(self, ssids, error_message):
        """Updates the GUI with PC Wi-Fi scan results. Runs on the main Tkinter thread."""
        page1 = self.frames["Page1_PCCheck"]
        page1.pc_wifi_listbox.delete(0, tk.END)

        if error_message:
            self.log_message(error_message)

        if ssids:
            self.log_message(f"Found {len(ssids)} PC Wi-Fi networks.")
            for i, ssid in enumerate(ssids):
                page1.pc_wifi_listbox.insert(tk.END, f"{i+1}. {ssid.strip()}")
            page1.connect_pc_wifi_button.config(state=tk.NORMAL)
        else:
            if not error_message:
                self.log_message("No PC Wi-Fi networks found.")
            page1.connect_pc_wifi_button.config(state=tk.DISABLED)

        page1.pc_wifi_refresh_button.config(state=tk.NORMAL)
        self.log_message("--- PC Wi-Fi scan finished. ---")

        self.get_and_display_current_ip_threaded_wrapper()


    def connect_pc_to_wifi(self, ssid, password):
        """
        Initiates Wi-Fi connection for the currently selected PC SSID.
        """
        self.log_message(f"\n--- Attempting to connect PC to '{ssid}' ---")
        self.log_message("Note: This feature might require administrator privileges.")

        page1 = self.frames["Page1_PCCheck"]
        page1.connect_pc_wifi_button.config(state=tk.DISABLED)
        page1.pc_wifi_refresh_button.config(state=tk.DISABLED)

        # Pass the selected SSID to the threaded function
        threading.Thread(target=self._connect_pc_to_wifi_threaded, args=(ssid, password), daemon=True).start()

    def _connect_pc_to_wifi_threaded(self, ssid, password):
        """Attempts to connect PC to a Wi-Fi network using netsh in a separate thread."""
        success_message = None
        error_message = None

        try:
            if password:
                # 1. Create XML profile
                profile_xml = self._create_wifi_profile_xml(ssid, password)
                profile_temp_path = os.path.join(os.environ["TEMP"], f"{ssid}.xml")

                with open(profile_temp_path, "w") as f:
                    f.write(profile_xml)
                self.log_message(f"Created temporary PC Wi-Fi profile XML: {profile_temp_path}")

                # 2. Add the profile
                add_profile_command = f'netsh wlan add profile filename="{profile_temp_path}" user=current'
                result_add = subprocess.run(
                    add_profile_command,
                    capture_output=True,
                    text=True,
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    shell=True
                )
                self.log_message(f"PC Profile added output:\n{result_add.stdout.strip()}")
                self.log_message(f"Successfully added PC Wi-Fi profile for '{ssid}'.")

                # Clean up the temporary XML file
                if os.path.exists(profile_temp_path):
                    os.remove(profile_temp_path)
                    self.log_message(f"Removed temporary PC Wi-Fi profile XML: {profile_temp_path}")

                # 3. Connect to the profile
                connect_command = f'netsh wlan connect name="{ssid}"'
                result_connect = subprocess.run(
                    connect_command,
                    capture_output=True,
                    text=True,
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    shell=True
                )
                success_message = f"Successfully sent connect command for PC to '{ssid}'. Check your system's Wi-Fi status.\n{result_connect.stdout.strip()}"
            else:
                # No password, use direct connect (for open networks or pre-existing profiles)
                connect_command = f'netsh wlan connect name="{ssid}"'
                result_connect = subprocess.run(
                    connect_command,
                    capture_output=True,
                    text=True,
                    check=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    shell=True
                )
                success_message = f"Successfully sent connect command for PC to '{ssid}'. Check your system's Wi-Fi status.\n{result_connect.stdout.strip()}"

        except subprocess.CalledProcessError as e:
            error_message = f"Error during PC Wi-Fi connection attempt: {e}\n  Command: {e.cmd}\n  Return Code: {e.returncode}\n  Stderr: {e.stderr.strip()}"
            if "Access is denied" in e.stderr:
                error_message += "\n--- Please ensure the script is run as ADMINISTRATOR! ---"
            elif "The specified network is not found" in e.stderr and password:
                error_message += "\n--- Ensure the password is correct or the network supports WPA2-PSK/AES. ---"
        except FileNotFoundError:
            error_message = "Error: 'netsh' command or temporary file path not found. This command is specific to Windows."
        except Exception as e:
            error_message = f"An unexpected error occurred during PC Wi-Fi connection: {e}"
            import traceback
            error_message += f"\n{traceback.format_exc()}"

        self.after(0, self._update_connect_pc_status_gui, success_message, error_message)

    def _update_connect_pc_status_gui(self, success_message, error_message):
        """Updates GUI after a PC connection attempt. Runs on main Tkinter thread."""
        page1 = self.frames["Page1_PCCheck"]
        if success_message:
            self.log_message(success_message)
        elif error_message:
            self.log_message(error_message)

        page1.connect_pc_wifi_button.config(state=tk.NORMAL)
        page1.pc_wifi_refresh_button.config(state=tk.NORMAL)
        self.log_message("--- PC Wi-Fi connection attempt finished. ---")

        self.get_and_display_current_ip_threaded_wrapper()


    def get_and_display_current_ip_threaded_wrapper(self):
        """Starts getting current PC IP and SSID in a separate thread."""
        self.current_ip_label.config(text="Scanning...")
        self.connected_ssid_label.config(text="Scanning...")
        threading.Thread(target=self._get_current_ip_threaded, daemon=True).start()

    def _get_current_ip_threaded(self):
        """
        Gets the current IP address and connected SSID of the specified Wi-Fi adapter.
        This runs in a background thread and updates GUI via self.after.
        """
        ip_address = "N/A"
        connected_ssid = "N/A"
        error_message = None

        try:
            # Command to get detailed network adapter information
            # Using chcp 65001 for UTF-8 encoding in subprocess output
            command = f"chcp 65001 && netsh interface ip show config name=\"{self.WIFI_ADAPTER_NAME}\""
            # self.log_message(f"DEBUG SYNC: _get_current_ip_threaded: Running netsh to get IP config for '{self.WIFI_ADAPTER_NAME}'...")
            result = subprocess.run(command, capture_output=True, text=True, check=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            ip_output = result.stdout
            # self.log_message(f"DEBUG SYNC: _get_current_ip_threaded: IP Config Output:\n{ip_output}")

            # Regex to find IP Address (v4)
            ip_match = re.search(r"IP Address:\s+((?:\d{1,3}\.){3}\d{1,3})", ip_output)
            if ip_match:
                ip_address = ip_match.group(1).strip()

            # Command to get WLAN status (for SSID)
            command_wlan = f"chcp 65001 && netsh wlan show interfaces"
            # self.log_message(f"DEBUG SYNC: _get_current_ip_threaded: Running netsh to get status for '{self.WIFI_ADAPTER_NAME}'...")
            result_wlan = subprocess.run(command_wlan, capture_output=True, text=True, check=True,
                                        creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            wlan_output = result_wlan.stdout
            # self.log_message(f"DEBUG SYNC: _get_current_ip_threaded: WLAN Status Output:\n{wlan_output}")

            # Regex to find SSID
            ssid_match = re.search(r"SSID\s+:\s+(.*)", wlan_output)
            if ssid_match:
                connected_ssid = ssid_match.group(1).strip()
            # self.log_message(f"DEBUG SYNC: _get_current_ip_threaded: Extracted SSID: '{connected_ssid}'")

        except subprocess.CalledProcessError as e:
            error_message = f"Error running netsh command for current IP/SSID: {e}\n  Stderr: {e.stderr.strip()}"
            # self.log_message(f"ERROR: {error_message}")
            if "No such interface is supported" in e.stderr or "The specified file was not found" in e.stderr:
                error_message += "\n--- Ensure your Wi-Fi Adapter Name is correct and exists on your system. ---"
            elif "Access is denied" in e.stderr:
                error_message += "\n--- Please ensure the script is run as ADMINISTRATOR! ---"
        except FileNotFoundError:
            error_message = "Error: 'netsh' command not found. This command is specific to Windows."
            # self.log_message(f"ERROR: {error_message}")
        except Exception as e:
            error_message = f"An unexpected error occurred getting current IP/SSID: {e}"
            import traceback
            error_message += f"\n{traceback.format_exc()}"
            # self.log_message(f"ERROR: {error_message}")

        self.after(0, self._update_ip_display_gui, ip_address, connected_ssid, error_message)


    def _update_ip_display_gui(self, ip_address, connected_ssid, error_message):
        """Updates the GUI labels with the current IP and SSID. Runs on main Tkinter thread."""
        self.current_ip_label.config(text=ip_address)
        self.connected_ssid_label.config(text=connected_ssid)
        if error_message:
            # Only show messagebox for critical errors
            if "No such interface is supported" in error_message or "not found" in error_message:
                messagebox.showerror("Error Getting PC Network Info", error_message)
            else:
                self.log_message(f"Warning: Issue getting PC network info: {error_message}")
        self.log_message(f"PC Network Status: IP={ip_address}, SSID='{connected_ssid}'")


    def set_static_ip_threaded_wrapper(self):
        """Starts the static IP setting process in a new thread."""
        threading.Thread(target=self._set_static_ip_threaded, daemon=True).start()

    def _set_static_ip_threaded(self):
        """
        Sets a static IP address for the PC's Wi-Fi adapter.
        Includes a check for JUICENET_SSID_PATTERN.
        """
        current_ip, connected_ssid, _ = self._get_current_ip_sync() # Get current status synchronously

        if not connected_ssid.startswith(self.JUICENET_SSID_PATTERN):
            warning_msg = (f"Warning: PC is currently connected to '{connected_ssid}'. "
                           f"Static IP changes are only allowed for JuiceNet devices (SSIDs starting with '{self.JUICENET_SSID_PATTERN}'). "
                           f"Aborting IP change.")
            self.log_message(warning_msg)
            self.after(0, lambda: messagebox.showwarning("IP Change Aborted", warning_msg))
            self.after(0, self.get_and_display_current_ip_threaded_wrapper) # Refresh display
            return
        else:
            self.log_message(f"DEBUG: Condition met: PC is connected to a JuiceNet device ('{connected_ssid}'). Proceeding with static IP change.")

        self.log_message(f"\n--- Attempting to set static IP for PC ({self.WIFI_ADAPTER_NAME})... ---")
        try:
            command = (
                f'netsh interface ip set address name="{self.WIFI_ADAPTER_NAME}" static '
                f'{self.TARGET_STATIC_IP} {self.TARGET_SUBNET_MASK} {self.TARGET_GATEWAY}'
            )
            self.log_message(f"Running command: {command}")
            result = subprocess.run(command, capture_output=True, text=True, check=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            self.log_message(result.stdout.strip())
            if result.stderr:
                self.log_message(f"Stderr from IP set command: {result.stderr.strip()}")

            if self.TARGET_DNS:
                dns_command = (
                    f'netsh interface ip set dns name="{self.WIFI_ADAPTER_NAME}" static {self.TARGET_DNS} primary'
                )
                self.log_message(f"Running DNS command: {dns_command}")
                result_dns = subprocess.run(dns_command, capture_output=True, text=True, check=True,
                                            creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
                self.log_message(result_dns.stdout.strip())
                if result_dns.stderr:
                    self.log_message(f"Stderr from DNS set command: {result_dns.stderr.strip()}")

            self.log_message(f"Successfully set PC's IP to static: {self.TARGET_STATIC_IP}")
            self.after(0, lambda: messagebox.showinfo("IP Set", f"PC IP successfully set to static {self.TARGET_STATIC_IP}"))
            self.ip_was_set_statically = True # Set the flag here!

        except subprocess.CalledProcessError as e:
            error_message = (f"Error setting PC static IP: {e}\n  Command: {e.cmd}\n  Return Code: {e.returncode}\n  "
                             f"Stderr: {e.stderr.strip()}")
            self.log_message(f"ERROR: {error_message}")
            if "Access is denied" in e.stderr:
                error_message += "\n--- Please ensure the script is run as ADMINISTRATOR! ---"
            self.after(0, lambda: messagebox.showerror("IP Set Error", error_message))
        except FileNotFoundError:
            error_message = "Error: 'netsh' command not found. This command is specific to Windows."
            self.log_message(f"ERROR: {error_message}")
            self.after(0, lambda: messagebox.showerror("IP Set Error", error_message))
        except Exception as e:
            error_message = f"An unexpected error occurred setting PC static IP: {e}"
            import traceback
            error_message += f"\n{traceback.format_exc()}"
            self.log_message(f"ERROR: {error_message}")
            self.after(0, lambda: messagebox.showerror("IP Set Error", error_message))
        finally:
            self.after(0, self.get_and_display_current_ip_threaded_wrapper) # Always refresh IP display

    def revert_ip_to_dhcp_threaded_wrapper(self):
        """Starts the DHCP IP reverting process in a new thread."""
        threading.Thread(target=self._revert_ip_to_dhcp_threaded, daemon=True).start()

    def _revert_ip_to_dhcp_threaded(self):
        """
        Reverts the PC's Wi-Fi adapter IP settings to DHCP.
        Includes a check for JUICENET_SSID_PATTERN.
        """
        current_ip, connected_ssid, _ = self._get_current_ip_sync() # Get current status synchronously

        if not connected_ssid.startswith(self.JUICENET_SSID_PATTERN):
            warning_msg = (f"Warning: PC is currently connected to '{connected_ssid}'. "
                           f"IP settings can only be reverted for JuiceNet devices (SSIDs starting with '{self.JUICENET_SSID_PATTERN}'). "
                           f"Aborting IP change.")
            self.log_message(warning_msg)
            self.after(0, lambda: messagebox.showwarning("IP Revert Aborted", warning_msg))
            self.after(0, self.get_and_display_current_ip_threaded_wrapper) # Refresh display
            return

        self.log_message(f"DEBUG: Condition met: PC is connected to a JuiceNet device ('{connected_ssid}'). Proceeding to revert IP to DHCP.")

        self.log_message(f"\n--- Attempting to revert PC IP to DHCP ({self.WIFI_ADAPTER_NAME})... ---")
        try:
            command_ip = f'netsh interface ip set address name="{self.WIFI_ADAPTER_NAME}" dhcp'
            self.log_message(f"Running command: {command_ip}")
            result_ip = subprocess.run(command_ip, capture_output=True, text=True, check=True,
                                       creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            self.log_message(f"IP DHCP command stdout: {result_ip.stdout.strip()}")
            if result_ip.stderr:
                self.log_message(f"IP DHCP command stderr: {result_ip.stderr.strip()}")

            command_dns = f'netsh interface ip set dns name="{self.WIFI_ADAPTER_NAME}" dhcp'
            self.log_message(f"Running DNS command: {command_dns}")
            result_dns = subprocess.run(command_dns, capture_output=True, text=True, check=True,
                                        creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            self.log_message(f"DNS DHCP command stdout: {result_dns.stdout.strip()}")
            if result_dns.stderr:
                self.log_message(f"DNS DHCP command stderr: {result_dns.stderr.strip()}")

            self.log_message("Successfully sent commands to revert PC IP to DHCP.")

            # --- VERIFY REVERSION ---
            # Give a moment for the system to process the change
            time.sleep(2)
            final_ip, final_ssid, verify_error = self._get_current_ip_sync()
            if verify_error:
                self.log_message(f"ERROR during post-revert IP verification: {verify_error}")
            else:
                self.log_message(f"Post-revert PC Network Status: IP={final_ip}, SSID='{final_ssid}'")
                # Basic check: if IP is still the static one, it likely failed.
                if final_ip == self.TARGET_STATIC_IP:
                    self.log_message("WARNING: PC IP still appears to be static after reversion attempt. Reversion likely failed.")
                    self.after(0, lambda: messagebox.showerror("IP Revert Warning", "PC IP still appears static. Reversion might have failed. Check logs."))
                else:
                    self.log_message("PC IP appears to have reverted from static, or was already DHCP.")
                    self.after(0, lambda: messagebox.showinfo("IP Reverted", "PC IP successfully reverted to DHCP."))

            self.ip_was_set_statically = False # Reset the flag

        except subprocess.CalledProcessError as e:
            error_message = (f"Error reverting PC IP to DHCP: {e}\n  Command: {e.cmd}\n  Return Code: {e.returncode}\n  "
                             f"Stderr: {e.stderr.strip()}")
            self.log_message(f"ERROR: {error_message}")
            if "Access is denied" in e.stderr:
                error_message += "\n--- Please ensure the script is run as ADMINISTRATOR! ---"
            self.after(0, lambda: messagebox.showerror("IP Revert Error", error_message))
        except FileNotFoundError:
            error_message = "Error: 'netsh' command not found. This command is specific to Windows."
            self.log_message(f"ERROR: {error_message}")
            self.after(0, lambda: messagebox.showerror("IP Revert Error", error_message))
        except Exception as e:
            error_message = f"An unexpected error occurred reverting PC IP to DHCP: {e}"
            import traceback
            error_message += f"\n{traceback.format_exc()}"
            self.log_message(f"ERROR: {error_message}")
            self.after(0, lambda: messagebox.showerror("IP Revert Error", error_message))
        finally:
            self.after(0, self.get_and_display_current_ip_threaded_wrapper)


    def _get_current_ip_sync(self):
        """
        Synchronously gets the current IP address and connected SSID of the specified Wi-Fi adapter.
        Used internally by IP setting functions.
        Returns: (ip_address, connected_ssid, error_message)
        """
        ip_address = "N/A"
        connected_ssid = "N/A"
        error_message = None

        try:
            # Command to get detailed network adapter information
            command = f"chcp 65001 && netsh interface ip show config name=\"{self.WIFI_ADAPTER_NAME}\""
            result = subprocess.run(command, capture_output=True, text=True, check=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            ip_output = result.stdout

            ip_match = re.search(r"IP Address:\s+((?:\d{1,3}\.){3}\d{1,3})", ip_output)
            if ip_match:
                ip_address = ip_match.group(1).strip()

            # Command to get WLAN status (for SSID)
            command_wlan = f"chcp 65001 && netsh wlan show interfaces"
            result_wlan = subprocess.run(command_wlan, capture_output=True, text=True, check=True,
                                        creationflags=subprocess.CREATE_NO_WINDOW, shell=True)
            wlan_output = result_wlan.stdout

            ssid_match = re.search(r"SSID\s+:\s+(.*)", wlan_output)
            if ssid_match:
                connected_ssid = ssid_match.group(1).strip()

        except subprocess.CalledProcessError as e:
            error_message = f"Error running netsh command synchronously for IP/SSID: {e}\n  Stderr: {e.stderr.strip()}"
        except FileNotFoundError:
            error_message = "Error: 'netsh' command not found. This command is specific to Windows."
        except Exception as e:
            error_message = f"An unexpected error occurred getting current IP/SSID synchronously: {e}"

        return ip_address, connected_ssid, error_message


    def start_automation(self, target_device_wifi_ssid, target_device_wifi_password):
        """Starts the web automation process in a separate thread."""
        if self.automation_thread and self.automation_thread.is_alive():
            messagebox.showwarning("Automation Running", "Automation is already in progress.")
            return

        page3 = self.frames["Page3_Automation"]
        page3.start_button.config(state=tk.DISABLED)
        # Disable all navigation buttons during automation
        for page_name in self.frames:
            page = self.frames[page_name]
            if hasattr(page, 'next_button') and page.next_button:
                page.next_button.config(state=tk.DISABLED)
            if hasattr(page, 'prev_button') and page.prev_button:
                 page.prev_button.config(state=tk.DISABLED)

        self.log_message("\n--- Starting web automation in background... ---")

        self.resume_automation_event.clear() # Clear any previous resume signal
        self.automation_finished_flag = False # Reset automation completion flag

        # Pass the instance method set_static_ip_threaded_wrapper as a callback
        self.automation_thread = threading.Thread(
            target=automate_web_actions,
            args=(
                self.TARGET_URL,
                self.EDGE_DRIVER_PATH,
                self.COMMAND_TO_EXECUTE,
                self.log_message,
                self.set_static_ip_threaded_wrapper, # Pass the callback here
                target_device_wifi_ssid,
                target_device_wifi_password,
                self.resume_automation_event, # Pass the threading.Event
                self # Pass the app instance to update GUI from thread
            ),
            daemon=True
        )
        self.automation_thread.start()
        self.after(100, self.check_automation_thread) # Start checking thread status

    def resume_automation(self):
        """Called when the 'Resume Script' button is clicked."""
        self.log_message("\n--- 'Resume Script' button clicked. ---")

        # Perform an immediate check on PC's connection status
        ip_address, connected_ssid, error_message = self._get_current_ip_sync()

        if error_message:
            self.log_message(f"ERROR during resume check: {error_message}")
            messagebox.showerror("Connection Check Error", f"Could not verify PC's connection: {error_message}. Please check logs.")
            return

        self.log_message(f"Current PC connection for resume check: IP={ip_address}, SSID='{connected_ssid}'")

        if connected_ssid.startswith(self.JUICENET_SSID_PATTERN):
            self.log_message("PC is connected to a JuiceNet SSID. Resuming automation.")
            self.resume_automation_event.set() # Signal the automation thread to continue
            # The disable_resume_button will be called by automation_thread.finally block via app_instance.after(0, app_instance.automation_finished_callback)
        else:
            self.log_message(f"WARNING: PC is currently connected to '{connected_ssid}'. "
                             f"Please manually connect to a network starting with '{self.JUICENET_SSID_PATTERN}' before resuming.")
            messagebox.showwarning("Not Connected to JuiceNet",
                                   f"Your PC is currently connected to '{connected_ssid}'. "
                                   f"Please connect to a JuiceNet network (e.g., JuiceNet-BC9) before clicking 'Resume'.")
            self.get_and_display_current_ip_threaded_wrapper() # Refresh PC IP display


    def automation_finished_callback(self):
        """Called by the automation thread when it finishes."""
        self.log_message("\n--- Automation thread has completed its work. ---")
        self.automation_finished_flag = True
        self.get_and_display_current_ip_threaded_wrapper() # Refresh PC IP display

        # Re-enable navigation buttons (especially the 'Next' to Page 4)
        for page_name in self.frames:
            page = self.frames[page_name]
            # Re-enable 'Back' buttons on all pages
            if hasattr(page, 'prev_button') and page.prev_button:
                 page.prev_button.config(state=tk.NORMAL)

        # Enable the specific "Next" button on Page 3 and re-enable start button
        page3 = self.frames["Page3_Automation"]
        if page3.next_button: # Check if it exists
            page3.next_button.config(state=tk.NORMAL)
        page3.start_button.config(state=tk.NORMAL) # Allow restart if desired
        page3.resume_script_button.config(state=tk.DISABLED) # Ensure resume is disabled


    def check_automation_thread(self):
        """Checks if the automation thread is still alive."""
        if self.automation_thread and self.automation_thread.is_alive():
            # Keep checking as long as it's alive. automation_finished_callback will handle UI updates.
            self.after(1000, self.check_automation_thread) # Check again in 1 second
        # If it's not alive, automation_finished_callback has already been called.


    def on_exit(self):
        """Handles graceful exit of the application when 'Exit' button is clicked."""
        if self.ip_was_set_statically:
            # Check current page and if it's not Page4_Cleanup
            if self.current_frame.__class__ != Page4_Cleanup:
                response = messagebox.askyesno(
                    "Confirm Exit - IP Not Reverted!",
                    "Your PC's IP was set to static by this program and has NOT been reverted to DHCP.\n"
                    "It is highly recommended to revert it before closing the application.\n\n"
                    "Do you want to go to the 'Cleanup' page to revert your IP now? "
                    "(Choosing 'No' will exit without reverting IP, potentially causing network issues.)"
                )
                if response:
                    self.show_frame(Page4_Cleanup)
                    return # Prevent immediate exit, let user go to cleanup page
                else:
                    self.log_message("WARNING: User chose to exit without reverting PC IP to DHCP. IP may remain static.")

        if self.automation_thread and self.automation_thread.is_alive():
            if messagebox.askyesno("Confirm Exit", "Automation is still running. Do you want to force quit?"):
                self.log_message("Force quitting application and trying to terminate automation thread.")
                self.resume_automation_event.set() # Unblock if it's waiting
                self.automation_thread.join(timeout=3) # Give it a moment to react
                self.destroy() # Close the Tkinter window
            else:
                return # User cancelled exit
        else:
            self.log_message("Exiting application...")
            self.destroy() # Close the Tkinter window directly


    def on_closing(self):
        """
        This method is called when the Tkinter window's X button is clicked.
        It prompts the user to revert IP if it was set by the program.
        """
        if self.ip_was_set_statically:
            response = messagebox.askyesno(
                "Confirm Exit - IP Not Reverted!",
                "Your PC's IP was set to static by this program and has NOT been reverted to DHCP.\n"
                "It is highly recommended to revert it before closing the application.\n\n"
                "Do you want to go to the 'Cleanup' page to revert your IP now? "
                "(Choosing 'No' will exit without reverting IP, potentially causing network issues.)"
            )
            if response:
                self.show_frame(Page4_Cleanup)
                return # Prevent closing, show cleanup page
            else:
                self.log_message("WARNING: User chose to close window without reverting PC IP to DHCP. IP may remain static.")

        if self.automation_thread and self.automation_thread.is_alive():
            # If automation is running, prompt for force quit
            if messagebox.askyesno("Confirm Exit", "Automation is still running. Do you want to force quit?"):
                self.log_message("Force quitting application and trying to terminate automation thread.")
                self.resume_automation_event.set() # Unblock if it's waiting
                self.automation_thread.join(timeout=3) # Give it a moment to react
                self.destroy()
            else:
                return # User cancelled closing
        else:
            self.log_message("Destroying Tkinter window.")
            self.destroy() # Close the Tkinter window

if __name__ == "__main__":
    app = Application()
    app.mainloop()
