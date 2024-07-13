# wordcloud_generator.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns
import pickle
import re
import matplotlib.font_manager as fm
from collections import Counter

class WordCloudApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WordCloud Generator")
        
        # 최소 크기 설정
        self.root.minsize(800, 600)
        
        # 기본 글꼴 설정
        default_font = ("맑은 고딕", 10)
        self.root.option_add("*Font", default_font)
        
        # matplotlib 한글 폰트 설정
        font_path = "C:/Windows/Fonts/malgun.ttf"
        font_prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
        
        self.filepath = ""
        self.data = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # File selection button
        self.file_button = tk.Button(self.root, text="Select Excel/CSV File", command=self.load_file)
        self.file_button.pack(pady=10)
        
        # Data preview
        self.preview_label = tk.Label(self.root, text="Data Preview (head 8):")
        self.preview_label.pack(pady=5)
        
        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.pack(pady=5)
        
        # Column selection
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
        
        # Save and Load DataFrame buttons
        self.save_button = tk.Button(self.root, text="Save DataFrame", command=self.save_dataframe)
        self.save_button.pack(pady=5)
        
        self.load_button = tk.Button(self.root, text="Load DataFrame", command=self.load_dataframe)
        self.load_button.pack(pady=5)
    
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
        for col in self.tree.get_children():
            self.tree.delete(col)
        
        self.tree["columns"] = list(self.data.columns)
        for column in self.data.columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=100)
        
        for _, row in self.data.head(8).iterrows():
            self.tree.insert("", "end", values=list(row))
    
    def load_columns(self):
        self.column_listbox.delete(0, tk.END)
        for column in self.data.columns:
            self.column_listbox.insert(tk.END, column)
    
    def preprocess_text(self, text):
        text = re.sub(r'\s+', '', text)  # 공백 제거
        words = re.findall(r'\w+', text)  # 단어 추출
        return words
    
    def generate_wordcloud(self):
        selected_column_index = self.column_listbox.curselection()
        if not selected_column_index:
            messagebox.showwarning("Warning", "Please select a column.")
            return
        
        column_name = self.column_listbox.get(selected_column_index)
        
        text = " ".join(self.data[column_name].dropna().astype(str).apply(lambda x: ' '.join(self.preprocess_text(x))).tolist())
        
        # 한글 폰트 경로 설정
        font_path = "C:/Windows/Fonts/malgun.ttf"
        
        # 불용어 설정
        stopwords = set(['가', '은', '는', '이', '의', '에', '를', '로', '와', '과', '도', '으로'])
        wordcloud = WordCloud(width=800, height=400, background_color='white', font_path=font_path, stopwords=stopwords).generate(text)
        
        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if save_path:
            wordcloud.to_file(save_path)
            messagebox.showinfo("Success", f"WordCloud saved to {save_path}")

    def show_histogram(self, top_n=20):
        selected_column_index = self.column_listbox.curselection()
        if not selected_column_index:
            messagebox.showwarning("Warning", "Please select a column.")
            return
        
        column_name = self.column_listbox.get(selected_column_index)
        
        all_words = []
        stopwords = set(['가', '은', '는', '이', '의', '에', '를', '로', '와', '과', '도', '으로'])
        
        self.data[column_name].dropna().astype(str).apply(lambda x: all_words.extend([word for word in self.preprocess_text(x) if word not in stopwords]))
        
        word_counts = Counter(all_words)
        top_words = word_counts.most_common(top_n)
        words, counts = zip(*top_words)
        
        plt.figure(figsize=(10, 6))
        sns.barplot(x=list(counts), y=list(words))
        plt.title(f'{column_name} 분포 (상위 {top_n}개 단어)')
        plt.xlabel('빈도수')
        plt.ylabel('단어')
        
        plt.show()

    def save_dataframe(self):
        if self.data is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".pickle", filetypes=[("Pickle files", "*.pickle")])
            if save_path:
                with open(save_path, "wb") as file:
                    pickle.dump(self.data, file)
                messagebox.showinfo("Success", f"DataFrame saved to {save_path}")
        else:
            messagebox.showwarning("Warning", "No data to save.")
    
    def load_dataframe(self):
        load_path = filedialog.askopenfilename(filetypes=[("Pickle files", "*.pickle")])
        if load_path:
            with open(load_path, "rb") as file:
                self.data = pickle.load(file)
            self.preview_data()
            self.load_columns()

if __name__ == "__main__":
    root = tk.Tk()
    app = WordCloudApp(root)
    root.mainloop()
