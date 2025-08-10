import pandas as pd
import streamlit as st
from groq import Groq
import os
from dotenv import load_dotenv
import glob
import time

# Load environment variables
load_dotenv()

class ExcelChatbot:
    def __init__(self, excel_file_path=None):
        self.df = None
        self.excel_file_path = excel_file_path
        self.api_key = os.getenv('GROQ_API_KEY')
        if not self.api_key:
            st.error("❌ GROQ_API_KEY not found. Please set it in your .env file.")
        self.client = Groq(api_key=self.api_key)
    
    def find_excel_files(self, folder_path=None):
        """Find Excel files in specified directory or current directory"""
        if folder_path and os.path.exists(folder_path):
            search_path = os.path.join(folder_path, "*.xlsx")
            search_path2 = os.path.join(folder_path, "*.xls")
            excel_files = glob.glob(search_path) + glob.glob(search_path2)
            return [(os.path.basename(f), f) for f in excel_files]
        else:
            excel_files = glob.glob("*.xlsx") + glob.glob("*.xls")
            return [(f, f) for f in excel_files]
    
    def load_data(self, file_path, sheet_name=None):
        """Load Excel data into pandas DataFrame"""
        try:
            if sheet_name:
                self.df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                self.df = pd.read_excel(file_path)
            self.excel_file_path = file_path
            return True
        except Exception as e:
            st.error(f"Error loading Excel file: {str(e)}")
            return False
    
    def get_sheet_names(self, file_path):
        """Get all sheet names from Excel file"""
        try:
            xl_file = pd.ExcelFile(file_path)
            return xl_file.sheet_names
        except Exception as e:
            st.error(f"Error reading Excel sheets: {str(e)}")
            return []
    
    def get_data_summary(self):
        """Get a summary of the data for context"""
        if self.df is None:
            return "No data loaded"
        
        summary = f"""
        Dataset Summary:
        - Total rows: {len(self.df)}
        - Total columns: {len(self.df.columns)}
        - Columns: {', '.join(self.df.columns.tolist())}
        
        Sample data (first 3 rows):
        {self.df.head(3).to_string()}
        """
        return summary
    
    def query_data_with_ai(self, user_question):
        """Use Groq to interpret the question and generate pandas code"""
        data_context = self.get_data_summary()
        system_prompt = f"""
        You are a data analyst helping to answer questions about an Excel dataset using pandas.
        Here's the dataset information:
        {data_context}
        The DataFrame is available as 'df'.
        Generate appropriate pandas code to answer the user's question.
        Return ONLY the pandas code, no explanations or markdown formatting.
        Make sure the code is safe and doesn't modify the original data.
        """
        try:
            response = self.client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_question}
                ]
            )
            generated_code = response.choices[0].message.content.strip()
            if generated_code.startswith("```"):
                generated_code = "\n".join(generated_code.split("\n")[1:-1])
            return generated_code
        except Exception as e:
            return f"Error generating query: {str(e)}"
    
    def execute_query(self, pandas_code):
        """Safely execute the pandas code"""
        try:
            safe_dict = {
                'df': self.df,
                'pd': pd
            }
            result = eval(pandas_code, {"__builtins__": {}}, safe_dict)
            return result
        except Exception as e:
            return f"Error executing query: {str(e)}"
    
    def chat(self, user_question):
        """Main chat function"""
        if self.df is None:
            return "Please load data first"
        
        pandas_code = self.query_data_with_ai(user_question)
        result = self.execute_query(pandas_code)
        return pandas_code, result

def main():
    st.title("📊 Excel Data Chatbot (Groq)")
    st.write("Connect to your OneDrive synced Excel files!")
    
    if 'chatbot' not in st.session_state:
        st.session_state.chatbot = ExcelChatbot()
    
    st.sidebar.title("📁 Folder Settings")
    
    default_paths = [
        os.path.expanduser("~/OneDrive"),
        os.path.expanduser("~/OneDrive - Personal"),
        os.path.expanduser("~/OneDrive - Business"),
        "C:/Users/" + os.getenv('USERNAME', '') + "/OneDrive" if os.name == 'nt' else None
    ]
    default_paths = [p for p in default_paths if p and os.path.exists(p)]
    
    folder_path = st.sidebar.text_input(
        "OneDrive Folder Path:", 
        value=default_paths[0] if default_paths else "",
        help="Enter the path to your OneDrive folder containing Excel files"
    )
    
    if default_paths:
        st.sidebar.write("**Common OneDrive locations:**")
        for path in default_paths:
            if st.sidebar.button(f"📂 {os.path.basename(path)}", key=path):
                st.session_state.folder_path = path
                st.rerun()
    
    if 'folder_path' in st.session_state:
        folder_path = st.session_state.folder_path
    
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("🔄 Refresh Files"):
            st.rerun()
    
    sheet_name = None
    if folder_path and os.path.exists(folder_path):
        st.success(f"✅ Connected to: {folder_path}")
        file_list = st.session_state.chatbot.find_excel_files(folder_path)
        if file_list:
            file_options = [f"{name} (Modified: {pd.to_datetime(os.path.getmtime(path), unit='s').strftime('%Y-%m-%d %H:%M')})" 
                          for name, path in file_list]
            selected_index = st.selectbox("Select Excel file:", range(len(file_options)), 
                                        format_func=lambda x: file_options[x])
            selected_file_name, selected_file_path = file_list[selected_index]
            sheet_names = st.session_state.chatbot.get_sheet_names(selected_file_path)
            if sheet_names:
                sheet_name = st.selectbox("Select sheet:", sheet_names)
            auto_refresh = st.checkbox("🔄 Auto-refresh data every 30 seconds")
            if auto_refresh:
                if 'last_modified' not in st.session_state:
                    st.session_state.last_modified = os.path.getmtime(selected_file_path)
                current_modified = os.path.getmtime(selected_file_path)
                if current_modified > st.session_state.last_modified:
                    st.info("📄 File updated! Reloading...")
                    st.session_state.last_modified = current_modified
                    if st.session_state.chatbot.load_data(selected_file_path, sheet_name=sheet_name):
                        st.session_state.current_file = selected_file_path
                time.sleep(30)
                st.rerun()
            if st.button("Load Data") or 'current_file' not in st.session_state or st.session_state.get('current_file') != selected_file_path:
                if st.session_state.chatbot.load_data(selected_file_path, sheet_name=sheet_name):
                    st.session_state.current_file = selected_file_path
                    st.success(f"✅ Loaded: {selected_file_name} | Sheet: {sheet_name} | Rows: {len(st.session_state.chatbot.df)} | Columns: {len(st.session_state.chatbot.df.columns)}")
                    if 'messages' in st.session_state:
                        st.session_state.messages = []
        else:
            st.warning("⚠️ No Excel files found in the selected OneDrive folder.")
            st.write(f"Looking in: {folder_path}")
    else:
        st.error("❌ OneDrive folder path not found or invalid.")
        return

    if st.session_state.chatbot.df is not None:
        with st.expander("📋 View Data Preview"):
            st.dataframe(st.session_state.chatbot.df.head(10))
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Columns:**")
                for col in st.session_state.chatbot.df.columns:
                    st.write(f"• {col}")
            with col2:
                st.write("**Data Types:**")
                for col, dtype in st.session_state.chatbot.df.dtypes.items():
                    st.write(f"• {col}: {dtype}")
        st.write("---")
        st.write("### 💬 Ask Questions About Your Data")
        st.write("**Quick Questions:**")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Most common value in first column"):
                first_col = st.session_state.chatbot.df.columns[0]
                st.session_state.quick_question = f"What is the most common value in {first_col}?"
        with col2:
            if st.button("Show column summary"):
                st.session_state.quick_question = "Show me a summary of all columns"
        with col3:
            if st.button("Count total rows"):
                st.session_state.quick_question = "How many total rows are in this data?"
        if 'messages' not in st.session_state:
            st.session_state.messages = []
        if 'quick_question' in st.session_state:
            prompt = st.session_state.quick_question
            del st.session_state.quick_question
            st.session_state.messages.append({"role": "user", "content": prompt})
            code, result = st.session_state.chatbot.chat(prompt)
            response = f"**Generated Code:**\n```python\n{code}\n```\n\n**Result:**\n{result}"
            st.session_state.messages.append({"role": "assistant", "content": response})
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                if message["role"] == "assistant":
                    st.markdown(message["content"])
                else:
                    st.write(message["content"])
        if prompt := st.chat_input("Ask about your data..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.write(prompt)
            with st.chat_message("assistant"):
                with st.spinner("Analyzing your question..."):
                    code, result = st.session_state.chatbot.chat(prompt)
                st.markdown("**Generated Code:**")
                st.code(code, language="python")
                st.markdown("**Result:**")
                st.write(result)
                response_text = f"**Generated Code:**\n```python\n{code}\n```\n\n**Result:**\n{result}"
            st.session_state.messages.append({"role": "assistant", "content": response_text})

if __name__ == "__main__":
    if not os.getenv('GROQ_API_KEY'):
        st.error("❌ Please set your GROQ_API_KEY in a .env file")
        st.write("Example: `GROQ_API_KEY=your_api_key_here`")
    else:
        main()
