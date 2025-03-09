from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.core.window import Window



class OutputPlotApp(App):
    def build(self):
        # Set the window background color to white
        Window.clearcolor = (1,1,1,1) # RGBA (1, 1, 1, 1) represents white

        # Create a BoxLayout with horizontal orientation
        layout = BoxLayout(orientation="vertical",padding=10, spacing=10)
        layout1 = BoxLayout(orientation="horizontal",padding=10, spacing=10)

        # Create 6 TextInput widgets
        self.text_inputs = [TextInput(hint_text = f"Input{i+1}", size_hint=(None,None), width=400, height=30) for i in range(15)]

        # Add each TextInput to the layout
        for text_input in self.text_inputs:
            layout.add_widget(text_input)

        # Create a submit button
        button1 = Button(text="Plot", size_hint=(None,None), width=100, height=30)
        button2 = Button(text="Plot and save", size_hint=(None,None), width=100, height=30)

        button1.bind(on_press=self.on_submit)
        layout1.add_widget(button1)
        layout1.add_widget(button2)
        layout.add_widget(layout1)

        return layout

    def on_submit(self, instance):
        # Handle the submit button click event
        input_values = [text_input.text for text_input in self.text_inputs]
        print("Submitted values : ", input_values)

if __name__ == "__main__":
        OutputPlotApp().run()