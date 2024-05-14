import pandas as pd
import random
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog
from PIL import Image, ImageTk
from datetime import datetime, timedelta
import json

input_csv_file = None

def select_input_file():
    global input_csv_file
    input_csv_file = filedialog.askopenfilename(title="Select Input CSV File")
    root1.destroy()
    create_second_interface()

def create_second_interface():

    root2 = tk.Tk()
    root2.title("Job Generation Tool")
    root2.geometry("700x800")

    num_variants_label = tk.Label(root2, text="Number of Product Variants:")
    num_variants_label.pack()
    num_variants_entry = tk.Entry(root2)
    num_variants_entry.pack()

    min_process_length_label = tk.Label(root2, text="Minimum Process Length:")
    min_process_length_label.pack()
    min_process_length_entry = tk.Entry(root2)
    min_process_length_entry.pack()

    num_jobs_label = tk.Label(root2, text="Number of Jobs:")
    num_jobs_label.pack()
    num_jobs_entry = tk.Entry(root2)
    num_jobs_entry.pack()

    def generate_output_excel():
        num_product_variants = int(num_variants_entry.get())
        min_process_length = int(min_process_length_entry.get())
        num_jobs = int(num_jobs_entry.get())
        df_input = pd.read_csv(input_csv_file)

        variant_names = []
        variant_priorities = []
        for i in range(1, num_product_variants + 1):
            variant_name = simpledialog.askstring("Variant Name", f"Enter Variant Name {i}:")
            variant_priority = simpledialog.askinteger("Variant Priority", f"Enter Priority for {variant_name}:")
            variant_names.append(variant_name)
            variant_priorities.append(variant_priority)

        random_job_ids = random.sample(range(10000, 10000 + num_jobs), num_jobs)
        variant_sequence = [i + 1 for i in range(num_product_variants)] * (num_jobs // num_product_variants)
        remaining_variants = num_jobs % num_product_variants
        for i in range(remaining_variants):
            variant_sequence.append(i + 1)
        random.shuffle(variant_sequence)

        def generate_random_process_sequence():
            processes = df_input['Processes'].tolist()
            process_length = random.randint(min_process_length, len(processes))
            random.shuffle(processes)
            return ";".join(processes[:process_length])

        random_sequences = [generate_random_process_sequence() for _ in range(num_product_variants)]
        data = []
        json_data = {"steps": []}
        current_timestamp = datetime.now()
        time_increment = timedelta(minutes=30)

        for job_id, variant_idx in zip(random_job_ids, variant_sequence):
            variant_name = variant_names[variant_idx - 1]
            processes = random_sequences[variant_idx - 1].split(";")
            for process in processes:
                processing_time = random.randint(5, 20)
                setup_time = random.randint(40, 100)
                start_time = current_timestamp.strftime('%Y-%m-%d-%H-%M-%S')
                end_time = (current_timestamp + timedelta(minutes=processing_time + setup_time + 10)).strftime('%Y-%m-%d-%H-%M-%S')
                sequence = f"{process},milling,{processing_time},{setup_time},{start_time},{end_time}"
                quantity = random.randint(1, 10)
                release_date = (datetime.now() + timedelta(days=random.randint(1, 30))).strftime('%Y-%m-%d')
                completion_date = (datetime.now() + timedelta(days=random.randint(31, 60))).strftime('%Y-%m-%d')
                priority = variant_priorities[variant_idx - 1]
                data.append((job_id, variant_name, quantity, release_date, completion_date, priority, sequence))

                json_data["steps"].append({
                    "process": random.choice(["drilling", "milling", "turning"]),
                    "duration": processing_time,
                    "setup_time": setup_time,
                    "start_time": start_time,
                    "end_time": end_time
                })


                current_timestamp += time_increment

        columns = ['Job ID', 'Product Variant', 'Quantity', 'Release Date', 'Completion Date', 'Priority', 'Process Sequence']
        df_output = pd.DataFrame(data, columns=columns)
        output_file_path = filedialog.asksaveasfilename(title="Save Output Excel File", defaultextension=".xlsx",
                                                        filetypes=[("Excel files", "*.xlsx")])
        df_output.to_excel(output_file_path, index=False)

        json_output_path = filedialog.asksaveasfilename(title="Save Output JSON File", defaultextension=".json", filetypes=[("JSON files", "*.json")])
        with open(json_output_path, 'w') as json_file:
            json.dump(json_data, json_file, indent=4)

        root2.destroy()

    generate_button = ttk.Button(root2, text="Generate Output Excel & JSON", command=generate_output_excel)
    generate_button.pack(pady=10)

    exit_button = ttk.Button(root2, text="Exit", command=root2.quit)
    exit_button.pack(pady=10)

    root2.mainloop()

root1 = tk.Tk()
root1.title("Select Process Plan File")
root1.geometry("400x400")

logo_image = Image.open(r"c:\Users\JLM-MS\Desktop\Hiwi\Task 1\1920px-Fraunhofer-Gesellschaft_2009_logo.png")
logo_image = logo_image.resize((150, 80), Image.ANTIALIAS)
logo = ImageTk.PhotoImage(logo_image)

logo_label = tk.Label(root1, image=logo)
logo_label.pack(pady=20)

select_button = ttk.Button(root1, text="Select Process Plan File", command=select_input_file)
select_button.pack(pady=10)

root1.mainloop()
