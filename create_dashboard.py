import tkinter as tk
from tkinter import messagebox
import pandas as pd
import webbrowser
import tempfile
import json

# ================= LOAD DATA =================
def load_data():
    try:
        df = pd.read_csv(r"C:\Users\rksol\OneDrive\Desktop\AJAY_\youtube_channels.csv")

        # Clean column names
        df.columns = df.columns.str.strip()

        # Remove rows with missing values
        df = df.dropna(subset=["Category", "Language", "Name", "Subscribers (Millions)"])

        # Remove extra spaces
        df["Category"] = df["Category"].str.strip()
        df["Language"] = df["Language"].str.strip()
        df["Name"] = df["Name"].str.strip()

        # Fix capitalization
        df["Category"] = df["Category"].str.title()
        df["Language"] = df["Language"].str.title()

        # Remove duplicates
        df = df.drop_duplicates()

        return df

    except Exception as e:
        messagebox.showerror("Error", f"File not loading:\n{e}")
        return None


# ================= DASHBOARD =================
def create_dashboard(filtered_df, category, language):

    if filtered_df.empty:
        messagebox.showinfo("No Data", "No Match Found 😢")
        return

    top5 = filtered_df.sort_values(by="Subscribers (Millions)", ascending=False).head(5)

    names = list(top5["Name"])
    subs = list(top5["Subscribers (Millions)"])

    names_js = json.dumps(names)
    subs_js = json.dumps(subs)

    table_rows = ""
    for _, row in top5.iterrows():
        table_rows += f"""
        <tr>
            <td>{row['Name']}</td>
            <td>{row['Category']}</td>
            <td>{row['Language']}</td>
            <td>{row['Subscribers (Millions)']} M</td>
        </tr>
        """

    html = f"""
    <html>
    <head>
        <title>YouTube Dashboard</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

        <style>
            body {{
                margin: 0;
                font-family: Arial;
                background: #0f0f0f;
                color: white;
                text-align: center;
            }}

            .container {{
                max-width: 1000px;
                margin: auto;
                padding: 20px;
            }}

            .card {{
                background: #1a1a1a;
                padding: 20px;
                border-radius: 12px;
                margin-top: 20px;
            }}

            canvas {{
                max-height: 400px;
            }}

            table {{
                width: 100%;
                border-collapse: collapse;
            }}

            th, td {{
                padding: 10px;
                border-bottom: 1px solid #333;
            }}

            th {{
                background: red;
            }}
        </style>
    </head>

    <body>
        <div class="container">
            <h1>🎯 Top 5 Channels</h1>
            <h3>Category: {category} | Language: {language}</h3>

            <div class="card">
                <canvas id="chart"></canvas>
            </div>

            <div class="card">
                <table>
                    <tr>
                        <th>Name</th>
                        <th>Category</th>
                        <th>Language</th>
                        <th>Subscribers</th>
                    </tr>
                    {table_rows}
                </table>
            </div>
        </div>

        <script>
        const ctx = document.getElementById('chart');

        new Chart(ctx, {{
            type: 'bar',
            data: {{
                labels: {names_js},
                datasets: [{{
                    label: 'Subscribers (Millions)',
                    data: {subs_js},
                    backgroundColor: 'red'
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{
                        labels: {{ color: 'white' }}
                    }}
                }},
                scales: {{
                    x: {{ ticks: {{ color: 'white' }} }},
                    y: {{ ticks: {{ color: 'white' }} }}
                }}
            }}
        }});
        </script>
    </body>
    </html>
    """

    file = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
    file.write(html.encode("utf-8"))
    file.close()

    webbrowser.open(file.name)


# ================= FILTER FUNCTION =================
def apply_filter():
    category = category_input.get()
    language = language_input.get()

    filtered_df = df[
        (df["Category"] == category) &
        (df["Language"] == language)
    ]

    create_dashboard(filtered_df, category, language)


# ================= GUI =================
df = load_data()

if df is not None:

    root = tk.Tk()
    root.title("YouTube Dashboard")
    root.geometry("400x300")
    root.configure(bg="#111")

    # Clean dropdown values
    categories = sorted(df["Category"].dropna().unique())
    languages = sorted(df["Language"].dropna().unique())

    tk.Label(root, text="Select Category", fg="white", bg="#111").pack(pady=5)
    category_input = tk.StringVar()
    category_menu = tk.OptionMenu(root, category_input, *categories)
    category_menu.pack()

    tk.Label(root, text="Select Language", fg="white", bg="#111").pack(pady=5)
    language_input = tk.StringVar()
    language_menu = tk.OptionMenu(root, language_input, *languages)
    language_menu.pack()

    # Default values
    if categories:
        category_input.set(categories[0])
    if languages:
        language_input.set(languages[0])

    tk.Button(root, text="Show Dashboard", bg="red", fg="white", command=apply_filter).pack(pady=20)

    root.mainloop()