import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import scrolledtext
import matplotlib
matplotlib.use('TkAgg')

def render_latex_preview(text):
    """Render a preview of the email text with basic LaTeX support"""
    window = tk.Toplevel()
    window.title("Email Preview")
    window.geometry("600x400")
    
    # Create a scrollable text widget for preview
    preview_text = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=70, height=20)
    preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Insert the text (basic LaTeX rendering - just display as text for now)
    preview_text.insert(tk.END, "Email Preview:\n\n" + text)
    preview_text.config(state=tk.DISABLED)  # Make it read-only
    
    # Add close button
    close_btn = tk.Button(window, text="Close", command=window.destroy)
    close_btn.pack(pady=5)
    
    # Center the window
    window.transient()
    window.grab_set()
    
    # Don't use mainloop() here as it blocks the main application
