from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.button import MDRaisedButton, MDFlatButton
from kivymd.uix.gridlayout import MDGridLayout
from kivy.uix.scrollview import ScrollView
from kivy.properties import ObjectProperty
import json
from RFPlot import myplotext
import os
from kivymd.uix.dialog import MDDialog

KV = """
MDFloatLayout:
    MDBoxLayout:
        orientation: 'vertical'
        spacing: dp(8)
        padding: dp(8)

        MDTopAppBar:
            id: app_bar  # ✅ Add this ID
            title: "Test Case Manager"
            height: dp(15)  # Reduced height
            font_size: "12sp"
            specific_text_color: 0, 0, 0, 1  # Darker text
            left_action_items: [["play", lambda x: app.run_test()]]
            right_action_items: [["folder", lambda x: app.show_file_chooser()]]

        BoxLayout:
            orientation: "horizontal"
            spacing: dp(10)

            MDCard:
                size_hint: (0.25, 1)  # Left panel for test cases
                padding: dp(10)
                md_bg_color: 0.9, 0.9, 0.9, 1  # Light gray background

                MDScrollView:
                    size_hint: (1, 1)
                    do_scroll_x: False
                    bar_width: dp(8)  # ✅ Makes scrollbar visible
                    scroll_type: ['bars', 'content']  # ✅ Enables scrollbar interaction

                    MDGridLayout:
                        id: test_case_layout
                        cols: 1
                        spacing: dp(4)
                        adaptive_height: True  # Important for scroll
                        size_hint_y: None  # Allow dynamic height

            MDCard:
                size_hint: (0.7, 1)  # Right panel for details
                padding: dp(2)
                spacing: dp(5)
                md_bg_color: 1, 1, 1, 1  # White background

                MDBoxLayout:
                    orientation: "vertical"
                    spacing: dp(8)

                    # Message when no test case file is loaded
                    MDLabel:
                        id: no_test_case_label
                        text: "No test case file loaded"
                        halign: "center"
                        theme_text_color: "Secondary"
                        opacity: 1  # Set to 1 initially
                        size_hint_y: None
                        height: self.texture_size[1] if self.opacity == 1 else 0  # Ensure it collapses when hidden

                    # Hide input fields initially
                    MDScrollView:
                        size_hint: (1, 1)  # Make it take full height
                        do_scroll_x: False  # Disable horizontal scrolling
                        bar_width: dp(8)  # ✅ Makes scrollbar visible
                        scroll_type: ['bars', 'content']  # ✅ Enables scrollbar interaction

                        MDBoxLayout:
                            id: input_fields_container
                            orientation: "vertical"
                            spacing: dp(8)
                            size_hint_y: None
                            height: self.minimum_height  # Adjust height dynamically
                            opacity: 0
                            disabled: True

                            MDTextField:
                                id: files_path
                                hint_text: "Files Path"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)  # Ensuring proper height for input fields

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: sheet_name
                                    hint_text: "Sheet Name"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: file_len
                                    hint_text: "Label Length"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: tx_power
                                    hint_text: "Tx_Power"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDTextField:
                                id: section_keys
                                hint_text: "Section Keys"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: linex
                                    hint_text: "Line X"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: liney
                                    hint_text: "Line Y"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDTextField:
                                id: limit_label
                                hint_text: "Limit Label"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: x_key
                                    hint_text: "X Key"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: y_key
                                    hint_text: "Y Key"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: extra_key
                                    hint_text: "Extra Key"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: plot_col
                                    hint_text: "Plot Column"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: xrange
                                    hint_text: "X Range"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: yrange
                                    hint_text: "Y Range"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDBoxLayout:
                                orientation: "horizontal"
                                spacing: dp(10)
                                size_hint_y: None
                                height: dp(55)

                                MDTextField:
                                    id: xy
                                    hint_text: "XY"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: x
                                    hint_text: "X"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: y
                                    hint_text: "Y"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)

                            MDTextField:
                                id: title_name
                                hint_text: "Title"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)

                            MDTextField:
                                id: file_name
                                hint_text: "File Name Filter"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)
"""


class TestCaseApp(MDApp):
    test_cases = ObjectProperty({})  # Stores test cases
    file_manager = None
    test_case_states = {}

    def build(self):
        return Builder.load_string(KV)

    def show_file_chooser(self):
        """Opens file chooser to select JSON file."""
        self.file_manager = MDFileManager(
            exit_manager=self.exit_file_manager,
            select_path=self.on_file_selected,
            preview=False  # Disable file previews
        )
        self.file_manager.show(os.getcwd())

    def on_file_selected(self, file_path):
        """Handles file selection from file chooser."""
        if file_path.endswith(".json"):
            self.load_test_cases(file_path)
        self.exit_file_manager()

    def exit_file_manager(self, *args):
        """Closes the file manager."""
        if self.file_manager:
            self.file_manager.close()

    def load_test_cases(self, file_path="test_cases.json"):
        """Loads test cases from a JSON file and updates buttons dynamically."""
        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                self.test_cases = json.load(file)
        else:
            self.test_cases = {}

            # ✅ Update the App Bar title to show the loaded file name
        file_name = os.path.basename(file_path)  # Extract file name
        self.root.ids.app_bar.title = f"Loaded: {file_name}"  # Update title

        print(f"📂 Loaded test cases from: {file_path}")  # Debug log

        # Clear previous test cases
        test_case_layout = self.root.ids.test_case_layout
        test_case_layout.clear_widgets()

        if not self.test_cases:
            self.root.ids.no_test_case_label.opacity = 1
            self.root.ids.no_test_case_label.height = self.root.ids.no_test_case_label.texture_size[1]
            self.root.ids.input_fields_container.opacity = 0
            self.root.ids.input_fields_container.disabled = True
        else:
            self.root.ids.no_test_case_label.opacity = 0
            self.root.ids.no_test_case_label.height = 0  # This removes the space
            self.root.ids.input_fields_container.opacity = 1
            self.root.ids.input_fields_container.disabled = False

        # Populate buttons dynamically
        for test_name, details in self.test_cases.items():
            # Load stored status, default to "Disabled"
            status = details.get("status", "Disabled")
            self.test_case_states[test_name] = status

            # Create a horizontal layout for the test case + toggle button
            box = MDBoxLayout(size_hint_y=None, height=48)
            button = MDRaisedButton(
                text=test_name,
                size_hint_x=1,  # Make buttons expand to full width
                on_release=self.load_test_details
            )
            # test_case_layout.add_widget(button)

            # Toggle button (Enable/Disable)
            toggle_button = MDFlatButton(
                text=status,
                size_hint_x=0.3,
                on_release=self.toggle_test_case,
                on_double_tap=self.toggle_test_case,  # Double-click handler
            )

            # Store reference in button
            toggle_button.test_name = test_name

            # Set initial colors
            if status == "Enabled":
                toggle_button.theme_text_color = "Custom"
                toggle_button.text_color = (0, 0.6, 0, 1)  # Green
            else:
                toggle_button.theme_text_color = "Custom"
                toggle_button.text_color = (0.5, 0.5, 0.5, 1)  # Gray

            # Add to layout
            box.add_widget(button)
            box.add_widget(toggle_button)
            test_case_layout.add_widget(box)

        # Update grid layout height
        test_case_layout.height = len(self.test_cases) * 48  # Each button ~48dp high

    def toggle_test_case(self, toggle_button):
        """Toggle test case state between Enabled and Disabled."""
        test_name = toggle_button.test_name
        new_status = "Enabled" if self.test_case_states.get(test_name) == "Disabled" else "Disabled"

        # ✅ Update both dictionary and button text
        self.test_case_states[test_name] = new_status
        self.test_cases[test_name]["status"] = new_status  # ✅ Update JSON structure

        toggle_button.text = new_status
        toggle_button.text_color = (0, 0.6, 0, 1) if new_status == "Enabled" else (0.5, 0.5, 0.5, 1)

        # ✅ Save changes to JSON
        try:
            with open("test_cases.json", "w") as file:
                json.dump(self.test_cases, file, indent=4)
            print("✅ Test case status updated successfully!")
        except Exception as e:
            print(f"❌ Error saving JSON file: {e}")

    def load_test_details(self, button):
        """Loads test case details into input fields when a button is clicked."""
        self.selected_test_case = button.text  # ✅ Store selected test case
        test_data = self.test_cases.get(self.selected_test_case, {})

        # Hide "No test case file loaded" and show input fields
        self.root.ids.no_test_case_label.opacity = 0
        self.root.ids.input_fields_container.opacity = 1
        self.root.ids.input_fields_container.disabled = False

        for field_id, value in test_data.items():
            if field_id in self.root.ids:
                text_field = self.root.ids[field_id]

                # ✅ Bind event listeners to update JSON file on text change
                text_field.unbind(text=self.update_test_case_data)  # ✅ Remove previous bindings
                text_field.text = str(value)
                text_field.bind(text=self.update_test_case_data)  # ✅ Add fresh binding

    def update_test_case_data(self, instance, value):
        """Updates the JSON file when user modifies an input field."""
        if not hasattr(self, "selected_test_case") or self.selected_test_case not in self.test_cases:
            print("❌ No test case selected or test case not in dictionary")
            return  # No test case selected

        field_id = instance.id if instance.id else instance.hint_text.lower().replace(" ", "_")  # Fallback mapping
        new_value = instance.text.strip()  # Strip extra spaces

        # ✅ Debugging log
        print(f"Updating {self.selected_test_case}: {field_id} = {new_value}")

        # ✅ Ensure test case exists before modifying
        if self.selected_test_case in self.test_cases:
            self.test_cases[self.selected_test_case][field_id] = new_value
        else:
            print("❌ Selected test case not found in test_cases dictionary")
            return

        # ✅ Save updated data to JSON file
        try:
            with open("test_cases.json", "w") as file:
                json.dump(self.test_cases, file, indent=4)
            print("✅ JSON file updated successfully!")
        except Exception as e:
            print(f"❌ Error saving JSON file: {e}")

    def load_test_details_by_name(self, test_name):
        """Loads test case details into input fields by name."""
        test_data = self.test_cases.get(test_name, {})

        for field_id, value in test_data.items():
            if field_id in self.root.ids:
                self.root.ids[field_id].text = str(value)

    def collect_input_fields(self):
        """Collects all input field data into a dictionary."""
        test_data = {
            "files_path": self.root.ids.files_path.text,
            "sheet_name": self.root.ids.sheet_name.text,
            "file_len": self.root.ids.file_len.text,
            "tx_power": self.root.ids.tx_power.text,
            "section_keys": self.root.ids.section_keys.text,
            "linex": self.root.ids.linex.text,
            "liney": self.root.ids.liney.text,
            "limit_label": self.root.ids.limit_label.text,
            "x_key": self.root.ids.x_key.text,
            "y_key": self.root.ids.y_key.text,
            "extra_key": self.root.ids.extra_key.text,
            "plot_col": self.root.ids.plot_col.text,
            "xrange": self.root.ids.xrange.text,
            "yrange": self.root.ids.yrange.text,
            "xy": self.root.ids.xy.text,
            "x": self.root.ids.x.text,
            "y": self.root.ids.y.text,
            "title_name": self.root.ids.title_name.text,
            "file_name": self.root.ids.file_name.text,
        }
        return test_data

    def run_test(self):
        """Runs all enabled test cases by collecting input field data and passing it to a function."""
        enabled_tests = [name for name, state in self.test_case_states.items() if state == "Enabled"]

        if not enabled_tests:
            print("❌ No enabled test cases to run.")
            return

        errors = []  # List to store failed test case names

        # ✅ Always collect fresh data for each test
        for test_name in enabled_tests:
            self.load_test_details_by_name(test_name)  # Load test case details
            test_data = self.collect_input_fields()  # ✅ Collect latest input field values
            print(f"▶️ Running test: {test_name}")
            try:
                self.run_test_case(test_name, test_data)  # ✅ Pass fresh test_data
            except Exception as e:
                print(f"❌ Error running test {test_name}: {e}")
                errors.append(test_name)  # Store failed test name

        # ✅ Show appropriate popup based on success or failure
        if errors:
            self.show_error_popup(f"Test failed: {', '.join(errors)}")
        else:
            self.show_success_popup()

    def show_success_popup(self):
        """Displays a popup when all plots are generated successfully."""
        self.dialog = MDDialog(
            title="Success",
            text="All plots generated successfully!",
            buttons=[
                MDFlatButton(text="OK", on_release=lambda x: self.dialog.dismiss())
            ],
        )
        self.dialog.open()

    def show_error_popup(self, message):
        """Displays an error popup if a test fails."""
        self.dialog = MDDialog(
            title="Error",
            text=message,
            buttons=[MDFlatButton(text="OK", on_release=lambda x: self.dialog.dismiss())],
        )
        self.dialog.open()

    def run_test_case(self, test_name, data):
        """Runs an individual test case safely."""
        try:
            print(f"Executing test case: {test_name}")

            # Use `.get()` to provide default values in case some fields are missing
            myplotext(
                file_path=data.get('files_path', ''),
                sheet_key=data.get('sheet_name', ''),
                section_key=data.get('section_keys', ''),
                linex=data.get('linex', ''),
                liney=data.get('liney', ''),
                plot_col=data.get('plot_col', ''),
                x_key=data.get('x_key', ''),
                y_key=data.get('y_key', ''),
                extra_key=data.get('extra_key', ''),
                xrange=data.get('xrange', ''),
                yrange=data.get('yrange', ''),
                limit_label=data.get('limit_label', ''),
                title=data.get('title_name', ''),
                xy=data.get('xy', ''),
                x=data.get('x', ''),
                y=data.get('y', ''),
                filelen=data.get('file_len', ''),
                file_name=data.get('file_name', ''),
                tx_power=data.get('tx_power', ''),
            )
        except KeyError as e:
            print(f"❌ Missing required key: {e}")
        except Exception as e:
            print(f"❌ Error running test case {test_name}: {e}")


if __name__ == "__main__":
    TestCaseApp().run()
