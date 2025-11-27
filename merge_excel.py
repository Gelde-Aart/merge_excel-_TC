import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def merge_excel_files():
    """
    Selecteert meerdere Excel-bestanden en voegt ze samen tot één bestand.
    """
    # Verberg het standaard Tkinter hoofdvenster
    root = tk.Tk()
    root.withdraw()

    print("Selecteer de Excel-bestanden die je wilt samenvoegen...")

    # Stap 1: Open bestandsselectie venster
    file_paths = filedialog.askopenfilenames(
        title="Selecteer Excel-bestanden",
        filetypes=[("Excel bestanden", "*.xlsx *.xls")]
    )

    if not file_paths:
        print("Geen bestanden geselecteerd. Script gestopt.")
        return

    print(f"{len(file_paths)} bestanden geselecteerd. Bezig met inlezen...")

    all_dataframes = []

    # Stap 2: Loop door elk bestand en lees het in
    for file in file_paths:
        try:
            # Lees het Excel-bestand in een DataFrame
            df = pd.read_excel(file)
            
            # Optioneel: Voeg een kolom toe zodat je weet uit welk bestand de rij komt
            # df['Bronbestand'] = os.path.basename(file)
            
            all_dataframes.append(df)
            print(f"✅ Ingelezen: {os.path.basename(file)} ({len(df)} rijen)")
            
        except Exception as e:
            print(f"❌ Fout bij lezen van {file}: {e}")

    if not all_dataframes:
        print("Er is geen data om samen te voegen.")
        return

    # Stap 3: Voeg alles samen (concateneren)
    print("Bezig met samenvoegen...")
    try:
        # ignore_index=True zorgt voor een nieuwe, doorlopende nummering van de rijen
        merged_df = pd.concat(all_dataframes, ignore_index=True)
    except Exception as e:
        messagebox.showerror("Fout", f"Fout bij samenvoegen: {e}")
        return

    # Stap 4: Vraag waar het resultaat opgeslagen moet worden
    save_path = filedialog.asksaveasfilename(
        title="Sla samengevoegd bestand op",
        defaultextension=".xlsx",
        filetypes=[("Excel bestand", "*.xlsx")],
        initialfile="samengevoegd.xlsx"
    )

    if save_path:
        try:
            # Sla op naar Excel (index=False zorgt dat je geen extra nummer-kolom krijgt)
            merged_df.to_excel(save_path, index=False)
            print(f"🎉 Succes! Bestand opgeslagen als: {save_path}")
            print(f"Totaal aantal rijen: {len(merged_df)}")
            messagebox.showinfo("Klaar", "De bestanden zijn succesvol samengevoegd!")
        except Exception as e:
            messagebox.showerror("Fout", f"Kon bestand niet opslaan: {e}")
    else:
        print("Opslaan geannuleerd.")

if __name__ == "__main__":
    merge_excel_files()