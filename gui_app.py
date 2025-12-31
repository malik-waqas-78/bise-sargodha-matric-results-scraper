import tkinter as tk

def main():
    # Create the main window
    window = tk.Tk()
    window.title("Greetings")
    window.geometry("300x150") # Set a reasonable size for the window

    # Create a label widget
    label = tk.Label(
        window,
        text="hello waqas",
        font=("Arial", 18),
        pady=20 # Add some padding
    )
    label.pack()

    # Create a button widget that closes the window
    exit_button = tk.Button(
        window,
        text="Exit",
        command=window.destroy, # The command to execute on click
        width=10,
        height=2
    )
    exit_button.pack(pady=10)

    # Start the GUI event loop
    window.mainloop()

if __name__ == "__main__":
    main()
