"""GUI module for Payment Terms QuickBooks Import application.

This module provides a tkinter interface for importing payment terms from Excel
to QuickBooks Desktop.
"""

import threading
import tkinter as tk
from tkinter import filedialog, ttk

from .excel_processor import process_payment_terms


def select_excel_file() -> str:
    """Open a file dialog to select an Excel file.

    Returns:
        str: Path to selected Excel file, or empty string if cancelled
    """
    return filedialog.askopenfilename(
        title="Select Excel File", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )


def process_payment_terms_background(
    file_path: str,
    status_label: tk.Label,
    process_button: tk.Button,
    results_text: tk.Text,
) -> None:
    """Process payment terms from selected Excel file and save to QuickBooks in background.

    Args:
        file_path: Path to the selected Excel file
        status_label: Label to show current status
        process_button: Button to re-enable when done
        results_text: Text widget to display results
    """

    def process_in_thread():
        try:
            process_button.config(state="disabled")
            status_label.config(text="Reading payment terms from Excel...")

            created_terms = process_payment_terms(file_path)

            status_label.config(text="Payment terms import complete!")

            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, "Payment Terms Import Results:\n")
            results_text.insert(tk.END, "=" * 40 + "\n\n")

            if created_terms:
                results_text.insert(
                    tk.END, f"Successfully imported {len(created_terms)} payment terms:\n\n"
                )
                for term_name in created_terms:
                    results_text.insert(tk.END, f"✓ {term_name}\n")
            else:
                results_text.insert(tk.END, "No payment terms were imported.\n")
                results_text.insert(tk.END, "Please check:\n")
                results_text.insert(tk.END, "- Selected Excel file has 'payment_terms' sheet\n")
                results_text.insert(tk.END, "- Sheet has 'Name' and 'ID' columns\n")
                results_text.insert(tk.END, "- QuickBooks Desktop is running\n")

        except Exception as e:
            status_label.config(text="Error occurred!")
            results_text.delete("1.0", tk.END)
            results_text.insert(tk.END, f"Error importing payment terms:\n{str(e)}\n\n")
            results_text.insert(tk.END, "Please ensure:\n")
            results_text.insert(
                tk.END, "- QuickBooks Desktop is running and a company file is open\n"
            )
            results_text.insert(
                tk.END, "- Excel file has 'payment_terms' sheet with Name/ID columns\n"
            )
            results_text.insert(tk.END, "- You have appropriate permissions in QuickBooks\n")
            results_text.insert(
                tk.END, "- QuickBooks allows external applications to access data\n"
            )

        finally:
            process_button.config(state="normal")

    threading.Thread(target=process_in_thread, daemon=True).start()


def create_main_window() -> tk.Tk:
    """Create the main application window.

    Returns:
        tk.Tk: The main window
    """
    root = tk.Tk()
    root.title("Payment Terms QuickBooks Import")
    root.geometry("600x500")
    root.resizable(True, True)
    return root


def run_app() -> None:
    """Run the main application."""
    root = create_main_window()

    # Title label
    title_label = tk.Label(
        root, text="Payment Terms QuickBooks Import", font=("Arial", 16, "bold"), fg="darkblue"
    )
    title_label.pack(pady=10)

    # Button frame
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Select Excel File button
    select_button = tk.Button(
        button_frame,
        text="Select Excel File",
        font=("Arial", 12),
        bg="lightblue",
        width=30,
        height=2,
    )
    select_button.pack(pady=5)

    # Status frame
    status_frame = tk.Frame(root)
    status_frame.pack(pady=10)

    status_label = tk.Label(
        status_frame, text="Select an Excel file to import payment terms", font=("Arial", 10)
    )
    status_label.pack()

    # Results frame
    results_frame = tk.Frame(root)
    results_frame.pack(pady=10, padx=20, fill="both", expand=True)

    results_label = tk.Label(results_frame, text="Results:", font=("Arial", 12, "bold"))
    results_label.pack(anchor="w")

    # Text widget with scrollbar
    text_frame = tk.Frame(results_frame)
    text_frame.pack(fill="both", expand=True)

    results_text = tk.Text(
        text_frame, height=15, width=60, font=("Courier", 10), wrap=tk.WORD, bg="white", fg="black"
    )

    scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=results_text.yview)
    results_text.configure(yscrollcommand=scrollbar.set)

    results_text.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configure button click event
    def on_select_file():
        file_path = select_excel_file()
        if file_path:
            process_payment_terms_background(file_path, status_label, select_button, results_text)

    select_button.config(command=on_select_file)

    # Start the application
    root.mainloop()
