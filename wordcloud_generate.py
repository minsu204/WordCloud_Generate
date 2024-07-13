# wordcloud_generate.py
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns

class WordCloudApp:
    def __init__(self, root) :
        self.root = root
        self.root.title("WordCloud Generator")

        self.filepath = ""
        self.data = None

        self.setup_ui()

    def setup_ui(self):
        # File Select Button
        self.file_button = tk.Button(self.root, text="Select Excel/CSV File", command=self.load_file)
        self.file_button.pack(pady=10)

        # Data Preview
        self.preview_label = tk.Label(self.root, text="Data Preview  (Top 8):")
        self.preview_label.pack(pady=5)

        self.text_preview = tk.Text(self.root, height=10, width=80)
        self.text_preview.pack(pady=5)

        # Column Selection
        self.column_label = tk.Label(self.root, text="Select Column for WordCloud:")
        self.column_label.pack(pady=5)

        self.column_listbox = tk.Listbox(self.root, selectmode=tk.SINGLE, height=5)
        self.column_listbox.pack(pady=5)

        # Generate WordCloud button
        self.generate_button = tk.Button(self.root, text="Generate WordCloud", command=self.generate_wordcloud)
        self.generate_button.pack(pady=10)
        
        # Show Histogram button
        self.histogram_button = tk.Button(self.root, text="Show Histogram", command=self.show_histogram)
        self.histogram_button.pack(pady=10)

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if not self.filepath:
            return
        
        try:
            if self.filepath.endswith('.csv'):
                self.data = pd.read_csv(self.filepath)
            else:
                self.data = pd.read_excel(self.filepath)

            self.preview_data()
            self.load_columns()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def preview_data(self):
        self.text_preview.delete(1.0, tk.END)
        self.text_preview.insert(tk.END, self.data.head(8).to_string(index=False))

    def load_columns(self):
        self.column_listbox.delete(0, tk.END)
        for column in self.data.columns:
            self.column_listbox.insert(tk.END, column)

    def generate_wordcloud(self):
        selected_column_index = self.column_listbox.curselection()
        if not selected_column_index:
            messagebox.showwarning("Warning", "Please select a column")
            return
        
        column_name = self.column_listbox.get(selected_column_index)

        text = " ".join(self.data[column_name].dropna().astype(str).tolist())
        wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)

        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if save_path:
            wordcloud.to_file(save_path)
            messagebox.showinfo("Success", f"WordCloud saved to {save_path}")

    def show_histogram(self):
        selected_column_index = self.column_listbox.curselection()
        if not selected_column_index:
            messagebox.showwarning("Warning", "Please select a column")
            return
        
        column_name = self.column_listbox.get(selected_column_index)

        plt.figure(figsize=(10, 6))
        sns.histplot(self.data[column_name].dropna(), kde=True)
        plt.title(f'Distribution of {column_name}')
        plt.xlabel(column_name)
        plt.ylabel('Frequency')

        plt.show()
    
# Main 
if __name__ == "__main__":
    root = tk.Tk()
    app = WordCloudApp(root)
    root.mainloop()