# ☕ Café Cultura — Excel Updater

A Streamlit web application for automatically updating training completion status in Excel files based on export data.

**© 2025 Rodrigo Bermudez — Café Cultura**

---

## 📋 What This App Does

The **Café Cultura Excel Updater** is a powerful tool designed to streamline the process of updating training records. Here's what it can do:

- **Upload two Excel files**: Your main training file and an export file containing completion status
- **Smart name matching**: Automatically matches users by combining first and last names (case-insensitive)
- **Selective updates**: Only marks trainings as "Completed" when the export file's Status column shows "Completed"
- **Flexible column selection**: Choose which training column to update
- **Visual highlighting**: Automatically highlights entire rows in yellow when ALL trainings in a defined range are completed
- **Easy download**: Download the updated and formatted Excel file with one click

### Key Features

✅ **User-friendly interface** with emoji icons and clean design
✅ **Customizable column mapping** for both main and export files
✅ **Configurable training range** for row highlighting
✅ **Real-time preview** of uploaded and processed data
✅ **Statistics dashboard** showing total records, updates, and highlighted rows
✅ **Error handling** with helpful messages

---

## 🚀 Running Locally

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Installation Steps

1. **Clone or download this repository**

2. **Navigate to the project directory**
   ```bash
   cd excel-updater-app
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Streamlit app**
   ```bash
   streamlit run update_excel_app.py
   ```

5. **Open your browser**
   - The app will automatically open at `http://localhost:8501`
   - If not, manually navigate to the URL shown in your terminal

---

## ☁️ Deploying to Streamlit Cloud

Streamlit Cloud offers **free hosting** for your Streamlit apps! Follow these steps to deploy:

### Step 1: Push to GitHub

First, initialize a Git repository and push your code to GitHub:

```bash
# Initialize Git repository
git init

# Add all files
git add .

# Create initial commit
git commit -m "Initial commit — Café Cultura Excel Updater"

# Rename branch to main
git branch -M main

# Add your GitHub repository as remote (replace with your username and repo name)
git remote add origin https://github.com/<YOUR_USERNAME>/excel-updater-app.git

# Push to GitHub
git push -u origin main
```

### Step 2: Deploy on Streamlit Cloud

1. **Go to [Streamlit Cloud](https://share.streamlit.io)**

2. **Sign in** with your GitHub account

3. **Click "New app"**

4. **Select your repository**
   - Repository: `<YOUR_USERNAME>/excel-updater-app`
   - Branch: `main`
   - Main file path: `update_excel_app.py`

5. **Click "Deploy"**

6. **Wait a few minutes** for your app to build and deploy

7. **Share your app URL** with anyone! It will look like:
   `https://share.streamlit.io/<YOUR_USERNAME>/excel-updater-app/main/update_excel_app.py`

### Tips for Streamlit Cloud

- Your app will automatically redeploy when you push changes to GitHub
- Free tier includes unlimited public apps
- Apps automatically sleep after inactivity and wake up when accessed
- Check the [Streamlit Cloud documentation](https://docs.streamlit.io/streamlit-community-cloud) for more details

---

## 📖 How to Use the App

### Step-by-Step Guide

1. **Upload Main File**
   - Click "Choose your main Excel file"
   - Select your Excel file containing training records

2. **Upload Export File**
   - Click "Choose your export Excel file"
   - Select your Excel file with completion status

3. **Configure Column Mappings**
   - Select First Name and Last Name columns for both files
   - Choose the Status column from the export file
   - Select which training column you want to update

4. **Set Highlighting Range**
   - Choose the start column of your training range
   - Choose the end column of your training range
   - Rows where ALL trainings in this range are "Completed" will be highlighted yellow

5. **Process Files**
   - Click the "🚀 Process Files" button
   - Review the statistics and preview

6. **Download Updated File**
   - Click "⬇️ Download Updated Excel File"
   - Your file will include updates and yellow highlighting

---

## 🛠️ Technical Details

### Dependencies

- **Streamlit**: Web application framework
- **Pandas**: Data manipulation and analysis
- **OpenPyXL**: Excel file reading and writing with formatting support

### How It Works

1. **File Upload**: Uses Streamlit's file uploader to accept Excel files
2. **Name Matching**: Combines first and last names, converts to lowercase for case-insensitive matching
3. **Status Filtering**: Identifies records with "Completed" status in the export file
4. **Update Logic**: Matches names and updates the selected training column
5. **Highlighting**: Checks all training columns in the defined range and applies yellow fill to complete rows
6. **Export**: Generates downloadable Excel file with formatting preserved

---

## 📝 File Structure

```
excel-updater-app/
│
├── update_excel_app.py    # Main Streamlit application
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

---

## 🤝 Support

For questions, issues, or suggestions, please contact:

**Rodrigo Bermudez**
© 2025 Café Cultura

---

## 📄 License

This project is © 2025 Rodrigo Bermudez — Café Cultura. All rights reserved.

---

**Enjoy your Excel automation! ☕**
