import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk

def render_latex_preview(text):
    # Crée une nouvelle fenêtre pour l'aperçu
    window = tk.Toplevel()
    window.title("Email Preview (LaTeX)")

    # Préparation du rendu matplotlib
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.text(0.05, 0.5, text, fontsize=12, verticalalignment='center', wrap=True)
    ax.axis('off')

    # Intégration dans tkinter
    canvas = FigureCanvasTkAgg(fig, master=window)
    canvas.get_tk_widget().pack()
    canvas.draw()

    window.mainloop()
