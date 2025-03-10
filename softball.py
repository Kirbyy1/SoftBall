import json
from collections import defaultdict
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import logging

# Configure logging
logging.basicConfig(
    filename='game_stats_processor.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def process_game_stats(json_files, output_file):
    logger.info(f"Starting processing for {len(json_files)} JSON files to {output_file}")
    player_stats = defaultdict(lambda: {
        'singles': 0, 'doubles': 0, 'triples_plus': 0,
        'grounders': 0, 'fly_balls': 0, 'bunts': 0,
        'pa': 0, 'so_k': 0, 'so_swing': 0, 'bb': 0,
        'first_name': '', 'last_name': '', 'number': ''
    })

    for json_file in json_files:
        try:
            with open(json_file, 'r') as f:
                data = json.load(f)
            logger.info(f"Successfully loaded {json_file}")
        except Exception as e:
            logger.error(f"Failed to load {json_file}: {str(e)}")
            return f"Error loading {json_file}: {str(e)}"

        player_lookup = {}
        for team_id, players in data['team_players'].items():
            for player in players:
                player_lookup[player['id']] = {
                    'first_name': player['first_name'],
                    'last_name': player['last_name'],
                    'number': player['number']
                }

        for play in data['plays']:
            play_type = play['name_template']['template'].lower()
            final_details = play['final_details']

            batter_id = None
            for detail in final_details:
                if '${' in detail['template']:
                    batter_id = detail['template'].split('${')[1].split('}')[0]
                    break

            if batter_id and batter_id in player_lookup:
                stats = player_stats[batter_id]
                stats['first_name'] = player_lookup[batter_id]['first_name']
                stats['last_name'] = player_lookup[batter_id]['last_name']
                stats['number'] = player_lookup[batter_id]['number']
                stats['pa'] += 1

                if 'single' in play_type:
                    stats['singles'] += 1
                    detail_text = final_details[0]['template'].lower()
                    if 'ground' in detail_text:
                        stats['grounders'] += 1
                    elif 'fly' in detail_text:
                        stats['fly_balls'] += 1
                    elif 'bunt' in detail_text:
                        stats['bunts'] += 1
                elif 'double' in play_type and 'double play' not in play_type:
                    stats['doubles'] += 1
                    detail_text = final_details[0]['template'].lower()
                    if 'ground' in detail_text:
                        stats['grounders'] += 1
                    elif 'fly' in detail_text:
                        stats['fly_balls'] += 1
                elif 'triple' in play_type:
                    stats['triples_plus'] += 1
                elif 'strikeout' in play_type:
                    detail_text = final_details[0]['template'].lower()
                    if 'looking' in detail_text:
                        stats['so_k'] += 1
                    elif 'swinging' in detail_text:
                        stats['so_swing'] += 1
                elif 'walk' in play_type:
                    stats['bb'] += 1
                elif 'ground out' in play_type:
                    stats['grounders'] += 1
                elif 'fly out' in play_type or 'pop out' in play_type:
                    stats['fly_balls'] += 1

    excel_data = []
    for player_id, stats in player_stats.items():
        if stats['pa'] > 0:
            player_name = f"{stats['first_name']} {stats['last_name'][0]} ({stats['number']})"
            excel_data.append({
                'Player (Number)': player_name,
                'Plate Appearances (PA)': stats['pa'],
                'Singles': stats['singles'],
                'Doubles': stats['doubles'],
                'Triples+': stats['triples_plus'],
                'Grounders': stats['grounders'],
                'Fly Balls': stats['fly_balls'],
                'Bunts': stats['bunts'],
                'Strikeouts Looking (SO K)': stats['so_k'],
                'Strikeouts Swinging (SO ꓘ)': stats['so_swing'],
                'Walks (BB)': stats['bb']
            })

    if excel_data:
        df = pd.DataFrame(excel_data)
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Player Stats', index=False)
                worksheet = writer.sheets['Player Stats']
                for column in worksheet.columns:
                    max_length = max(len(str(cell.value)) for cell in column if cell.value) + 2
                    worksheet.column_dimensions[column[0].column_letter].width = max_length
            logger.info(f"Successfully saved statistics to {output_file}")
            return f"Statistics successfully saved to '{output_file}'"
        except Exception as e:
            logger.error(f"Failed to save to Excel: {str(e)}")
            return f"Error saving to Excel: {str(e)}"
    else:
        logger.warning("No player statistics to save (no plate appearances found)")
        return "No player statistics to save (no plate appearances found)"


class JsonProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON Game Stats Processor")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        self.json_files = []

        # Main frame with padding
        self.main_frame = ttk.Frame(root, padding="15")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        ttk.Label(self.main_frame, text="Game Statistics Processor", font=("Helvetica", 16, "bold")).grid(
            row=0, column=0, columnspan=3, pady=(0, 15)
        )

        # JSON Files Section
        ttk.Label(self.main_frame, text="Selected JSON Files:", font=("Helvetica", 10, "bold")).grid(
            row=1, column=0, sticky=tk.W, pady=(0, 5)
        )
        self.file_listbox = tk.Listbox(self.main_frame, height=12, width=60, borderwidth=2, relief="groove")
        self.file_listbox.grid(row=2, column=0, columnspan=3, pady=5, padx=(0, 10))

        # Buttons frame
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        ttk.Button(self.button_frame, text="Add JSON Files", command=self.add_files, width=15).grid(
            row=0, column=0, padx=5
        )
        ttk.Button(self.button_frame, text="Remove Selected", command=self.remove_file, width=15).grid(
            row=0, column=1, padx=5
        )

        # Output filename section
        self.output_frame = ttk.LabelFrame(self.main_frame, text="Output Settings", padding="10")
        self.output_frame.grid(row=4, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))

        ttk.Label(self.output_frame, text="Excel Filename:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(self.output_frame, width=40)
        self.output_entry.grid(row=0, column=1, pady=5, padx=5)
        self.output_entry.insert(0, "player_statistics.xlsx")
        ttk.Label(self.output_frame, text=".xlsx will be added if omitted").grid(row=0, column=2, sticky=tk.W)

        # Process button
        self.process_button = ttk.Button(self.main_frame, text="Process Files", command=self.process_files, width=20)
        self.process_button.grid(row=5, column=0, columnspan=3, pady=15)

        # Status bar
        self.status_bar = ttk.Label(self.main_frame, text="No files added yet", relief="sunken", anchor=tk.W)
        self.status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        # Configure grid weights
        self.main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        logger.info("Application started")

    def add_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("JSON files", "*.json")])
        if file_paths:
            new_files = [fp for fp in file_paths if fp not in self.json_files]
            if new_files:
                self.json_files.extend(new_files)
                for file_path in new_files:
                    self.file_listbox.insert(tk.END, os.path.basename(file_path))
                    logger.info(f"Added file: {file_path}")
                self.update_status()

    def remove_file(self):
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            removed_file = self.json_files.pop(index)
            self.file_listbox.delete(index)
            self.update_status()
            logger.info(f"Removed file: {removed_file}")

    def update_status(self):
        if self.json_files:
            self.status_bar.config(text=f"{len(self.json_files)} file(s) added")
        else:
            self.status_bar.config(text="No files added yet")

    def process_files(self):
        if not self.json_files:
            messagebox.showwarning("Warning", "Please add at least one JSON file.")
            logger.warning("Process attempted with no files added")
            return

        output_file = self.output_entry.get().strip()
        if not output_file:
            messagebox.showwarning("Warning", "Please enter an output filename.")
            logger.warning("Process attempted with no output filename")
            return
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'

        self.process_button.config(state='disabled')
        self.status_bar.config(text="Processing...")
        self.root.update()

        result = process_game_stats(self.json_files, output_file)

        self.process_button.config(state='normal')
        if result:
            self.status_bar.config(text=result)
            if "successfully" in result:
                messagebox.showinfo("Success", result)
            else:
                messagebox.showerror("Error", result)


if __name__ == "__main__":
    root = tk.Tk()
    app = JsonProcessorApp(root)
    root.mainloop()
    logger.info("Application closed")
