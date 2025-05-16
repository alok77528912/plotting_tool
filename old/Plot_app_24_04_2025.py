import os
import textwrap
import matplotlib.pyplot as plt
import numpy as np
import datetime as dt
import pandas as pd

from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.button import MDRaisedButton, MDFlatButton
from kivy.properties import ObjectProperty
import json
from kivymd.uix.dialog import MDDialog
import ast

from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.textfield import MDTextField

from plot_main import myplotext

KV = """
MDFloatLayout:
    MDBoxLayout:
        orientation: 'vertical'
        spacing: dp(8)
        padding: dp(8)

        MDTopAppBar:
            id: app_bar  # ✅ Add this ID
            title: "No json file loaded. Please load a json file"
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
                padding: dp(4)
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
                                    hint_text: "File_len"
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
                                    id: line_x
                                    hint_text: "Line X"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: line_y
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
                                    id: plot_column
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
                                    id: x_range
                                    hint_text: "X Range"
                                    font_size: "14sp"
                                    mode: "rectangle"
                                    size_hint_y: None
                                    height: dp(55)
                                MDTextField:
                                    id: y_range
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
                                hint_text: "Title_Name"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)

                            MDTextField:
                                id: file_name
                                hint_text: "File_Name"
                                font_size: "14sp"
                                mode: "rectangle"
                                size_hint_y: None
                                height: dp(48)
"""


class TestCaseApp(MDApp):
    test_cases = ObjectProperty({})
    file_manager = None
    test_case_states = {}
    selected_button = None
    loaded_json_path = None  # New

    def build(self):
        return Builder.load_string(KV)

    def show_file_chooser(self):
        self.file_manager = MDFileManager(
            exit_manager=self.exit_file_manager,
            select_path=self.on_file_selected,
            preview=False
        )
        self.file_manager.show(os.getcwd())

    def on_file_selected(self, file_path):
        try:
            if file_path.endswith(".json"):
                self.load_test_cases(file_path)
                self.exit_file_manager()
        except Exception as e:
            print("Error in json file\n",e)

    def exit_file_manager(self, *args):
        if self.file_manager:
            self.file_manager.close()

    def on_right_click(self, instance, touch, test_name):
        if touch.button == 'right' and instance.collide_point(*touch.pos):
            menu_items = [
                {"text": "Rename", "on_release": lambda x=test_name: self.rename_test_case(x)},
                {"text": "Duplicate", "on_release": lambda x=test_name: self.duplicate_test_case(x)},
                {"text": "Delete", "on_release": lambda x=test_name: self.confirm_delete_test_case(x)},
            ]
            self.menu = MDDropdownMenu(
                caller=instance,
                items=menu_items,
                width_mult=3
            )
            self.menu.open()

    def rename_test_case(self, old_name):
        self.menu.dismiss()
        text_field = MDTextField(hint_text="New name", text=old_name)
        self.dialog = MDDialog(
            title="Rename Test Case",
            type="custom",
            content_cls=text_field,
            buttons=[
                MDFlatButton(text="Cancel", on_release=lambda x: self.dialog.dismiss()),
                MDFlatButton(
                    text="Rename",
                    on_release=lambda x: self.confirm_rename_test_case(old_name, text_field.text)
                ),
            ],
        )
        self.dialog.open()

    def confirm_rename_test_case(self, old_name, new_name):
        self.dialog.dismiss()
        if new_name and new_name != old_name and new_name not in self.test_cases:
            self.test_cases[new_name] = self.test_cases.pop(old_name)
            self.test_case_states[new_name] = self.test_case_states.pop(old_name)
            self.save_json()
            self.load_test_cases(self.loaded_json_path)

    def duplicate_test_case(self, name):
        self.menu.dismiss()
        new_name = f"{name}_copy"
        index = 1
        while new_name in self.test_cases:
            new_name = f"{name}_copy{index}"
            index += 1
        self.test_cases[new_name] = json.loads(json.dumps(self.test_cases[name]))  # Deep copy
        self.test_case_states[new_name] = "Disabled"
        self.save_json()
        self.load_test_cases(self.loaded_json_path)

    def confirm_delete_test_case(self, name):
        self.menu.dismiss()
        self.dialog = MDDialog(
            title="Delete Test Case",
            text=f"Are you sure you want to delete '{name}'?",
            buttons=[
                MDFlatButton(text="Cancel", on_release=lambda x: self.dialog.dismiss()),
                MDFlatButton(
                    text="Delete",
                    on_release=lambda x: self.delete_test_case(name)
                ),
            ],
        )
        self.dialog.open()

    def delete_test_case(self, name):
        self.dialog.dismiss()
        if name in self.test_cases:
            del self.test_cases[name]
            self.test_case_states.pop(name, None)
            self.save_json()
            self.load_test_cases(self.loaded_json_path)

    def load_test_cases(self, file_path):
        self.loaded_json_path = file_path  # Track loaded file path
        with open(file_path, "r") as file:
            self.test_cases = json.load(file)

        # Remove unknown keys
        for tc in self.test_cases.values():
            tc.pop('', None)

        self.root.ids.app_bar.title = f"Loaded: {file_path}"
        layout = self.root.ids.test_case_layout
        layout.clear_widgets()

        if not self.test_cases:
            self.root.ids.no_test_case_label.opacity = 1
            self.root.ids.input_fields_container.opacity = 0
            self.root.ids.input_fields_container.disabled = True
            return

        self.root.ids.no_test_case_label.opacity = 0
        self.root.ids.input_fields_container.opacity = 1
        self.root.ids.input_fields_container.disabled = False

        for test_name, details in self.test_cases.items():
            status = details.get("status", "Disabled")
            self.test_case_states[test_name] = status

            box = MDBoxLayout(size_hint_y=None, height=35)
            button = MDRaisedButton(
                text=test_name,
                size_hint_x=1,
                on_release=self.load_test_details
            )
            button.bind(on_touch_down=lambda instance, touch, name=test_name: self.on_right_click(instance, touch, name))
            toggle = MDFlatButton(
                text=status,
                size_hint_x=0.3,
                on_release=self.toggle_test_case
            )
            toggle.test_name = test_name
            toggle.theme_text_color = "Custom"
            toggle.text_color = (0, 0.6, 0, 1) if status == "Enabled" else (0.5, 0.5, 0.5, 1)

            box.add_widget(button)
            box.add_widget(toggle)
            layout.add_widget(box)

        layout.height = len(self.test_cases) * 48

    def toggle_test_case(self, toggle_button):
        test_name = toggle_button.test_name
        new_status = "Enabled" if self.test_case_states[test_name] == "Disabled" else "Disabled"
        self.test_case_states[test_name] = new_status
        self.test_cases[test_name]["status"] = new_status
        toggle_button.text = new_status
        toggle_button.text_color = (0, 0.6, 0, 1) if new_status == "Enabled" else (0.5, 0.5, 0.5, 1)
        self.save_json()

    def load_test_details(self, button):
        # Reset previous selection
        if hasattr(self, 'selected_button') and self.selected_button:
            self.selected_button.md_bg_color = [0.01, 0.65, 1, 1]  # Default white

        # Highlight current selection
        button.md_bg_color = [0.7, 0.85, 1, 1]  # Light blue or any color you prefer
        self.selected_button = button

        self.selected_test_case = button.text
        test_data = self.test_cases.get(self.selected_test_case, {})
        self.root.ids.no_test_case_label.opacity = 0
        self.root.ids.input_fields_container.opacity = 1
        self.root.ids.input_fields_container.disabled = False

        for field_id, value in test_data.items():
            if field_id in self.root.ids:
                text_field = self.root.ids[field_id]
                text_field.unbind(text=self.update_test_case_data)
                try:
                    parsed_value = ast.literal_eval(str(value))
                    text_field.text = str(parsed_value)
                except:
                    text_field.text = str(value)
                text_field.bind(text=self.update_test_case_data)

    def update_test_case_data(self, instance, value):
        if not hasattr(self, "selected_test_case") or self.selected_test_case not in self.test_cases:
            return

        field_id = instance.id if instance.id else instance.hint_text.lower().replace(" ", "_")
        raw_input = instance.text.strip()

        cast_fields = {
            "files_path":str,"sheet_name":str,"file_len":str,"tx_power":str,"section_keys":str,"line_x":str,
            "line_y":str,"limit_label":str,"x_key":str,"y_key":str,"extra_key":str,"plot_column":str,"x_range":str,
            "y_range":str,"xy":int,"x":int,"y":int,"title_name":str,"file_name":str
        }

        try:
            if field_id in cast_fields:
                parsed_value = cast_fields[field_id](raw_input)
            else:
                parsed_value = ast.literal_eval(raw_input)
        except:
            parsed_value = raw_input

        if field_id == "file_len":
            # Update file_len in all test cases
            for name in self.test_cases:
                self.test_cases[name]["file_len"] = parsed_value
        else:
            self.test_cases[self.selected_test_case][field_id] = parsed_value
        # self.test_cases[self.selected_test_case][field_id] = parsed_value

        self.save_json()

    def save_json(self):
        if self.loaded_json_path:
            with open(self.loaded_json_path, "w") as file:
                json.dump(self.test_cases, file, indent=4)
            print("✅ JSON updated successfully")
        else:
            print("❌ No loaded file to save to")

    def load_test_details_by_name(self, test_name):
        """Loads test case details into input fields by name."""
        test_data = self.test_cases.get(test_name, {})

        for field_id, value in test_data.items():
            if field_id in self.root.ids:
                self.root.ids[field_id].text = str(value)

    def collect_input_fields(self):
        """Collects all input field data into a dictionary."""
        return {
            "files_path": self.root.ids.files_path.text,
            "sheet_name": self.root.ids.sheet_name.text,
            "file_len": self.root.ids.file_len.text,
            "tx_power": self.root.ids.tx_power.text,
            "section_keys": self.root.ids.section_keys.text,
            "line_x": self.root.ids.line_x.text,
            "line_y": self.root.ids.line_y.text,
            "limit_label": self.root.ids.limit_label.text,
            "x_key": self.root.ids.x_key.text,
            "y_key": self.root.ids.y_key.text,
            "extra_key": self.root.ids.extra_key.text,
            "plot_column": self.root.ids.plot_column.text,
            "x_range": self.root.ids.x_range.text,
            "y_range": self.root.ids.y_range.text,
            "xy": self.root.ids.xy.text,
            "x": self.root.ids.x.text,
            "y": self.root.ids.y.text,
            "title_name": self.root.ids.title_name.text,
            "file_name": self.root.ids.file_name.text,
        }

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
                errors.append(f"{test_name} ({e})")  # Store failed test name

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

    def string_to_bool(self,s):
        if s == "1":
            return True
        elif s == "0" or s == "true" or s == "TRUE":
            return False
        else:
            return False

    def run_test_case(self, test_name, data):
        # try:
        print(f"Executing test case: {test_name}")

        # Use `.get()` to provide default values in case some fields are missing
        file_path = data.get('files_path', '')
        sheet_key = data.get('sheet_name', '')
        section_key = [item.strip() for item in data.get('section_keys', '').split(',') if item.strip()]
        linex = [[float(value) for value in group.split(':')] for group in data.get('line_x', '').strip().split(',') if len(data.get('line_x', '').strip()) > 0]
        liney = [[float(value) for value in group.split(':')] for group in data.get('line_y', '').strip().split(',') if len(data.get('line_y', '').strip()) > 0]
        plot_col = int(data.get('plot_column', ''))
        x_key = data.get('x_key', '')
        y_key = data.get('y_key', '')
        if "," in data.get('extra_key', ''):
            extra_key = [item.strip() for item in data.get('extra_key', '').split(',') if item.strip() if len(data.get('extra_key', '').strip()) > 0]
        else:
            extra_key = data.get('extra_key', '')
        if ":" in data.get('x_range', '').strip():
            xrange = [float(value) for value in data.get('x_range', '').strip().split(':')]
        else:
            xrange = []
        if ":" in data.get('y_range', '').strip():
            yrange = [float(value) for value in data.get('y_range', '').strip().split(':')]
        else:
            yrange = []
        if "," in data.get('limit_label', ''):
            limit_label = [item.strip() for item in data.get('limit_label', '').split(',') if item.strip()]
        else:
            limit_label = data.get('limit_label', '')
        title = data.get('title_name', '')
        xy = self.string_to_bool(data.get('xy', ''))
        x = self.string_to_bool(data.get('x', ''))
        y = self.string_to_bool(data.get('y', ''))
        filelen = int(data.get('file_len', ''))
        file_name = data.get('file_name', '')
        tx_power = float(data.get('tx_power', ''))

        myplotext(file_path=file_path, sheet_key=sheet_key, section_key=section_key, linex=linex, liney=liney, plot_col=plot_col, x_key=x_key, y_key=y_key,
                       extra_key=extra_key, xrange=xrange, yrange=yrange, limit_label=limit_label, title=title, xy=xy, x=x, y=y, filelen=filelen,file_name=file_name,tx_power=tx_power)
        # except KeyError as e:
        #     print(f"❌ Missing required key: {e}")
        # except Exception as e:
        #     print(f"❌ Error running test case {test_name}: {e}")
        # """Runs an individual test case safely."""

if __name__ == "__main__":
    TestCaseApp().run()
