import openpyxl
import sys
import csv
import os

import dearpygui.dearpygui as imgui

imgui.create_context()

def error_window(err_desc: str):
    with imgui.window(label="Virhe"):
        imgui.add_text(err_desc)

def is_revolut(filename):
    return os.path.splitext(filename)[1] == '.csv'

def kill(exitcode):
    imgui.destroy_context()
    exit(exitcode)

def load_revolut(filepath):
        creditAcct = imgui.get_value("__revolut_credit_account") # 2880
        topupAcct = imgui.get_value("__revolut_topup_account") # 1700
        debitAcct = imgui.get_value("__revolut_debit_account") # 1930

        data = []
        
        with open(filepath) as file:
            reader = csv.reader(file)
            next(reader)
            for line in reader:
                if len(line) < 3:
                    continue
                dataEntry = {}

                isTopup = line[3] == "TOPUP"

                dataEntry["name"] = line[5]
                dataEntry["amount"] = float(line[12])
                dataEntry["credit"] = debitAcct if isTopup else creditAcct
                dataEntry["debit"] = topupAcct if isTopup else debitAcct

                data.append(dataEntry)
            pass

        return data

def read_workbook(filepath):
    workbook = openpyxl.load_workbook(filename=filepath)

    if not workbook:
        kill(1)

    worksheet = workbook["in"]
    row_data = []
    for row in worksheet.iter_rows(values_only=True, max_col=6):
        row_data.append({
            "name": row[1],
            "credit": row[2],
            "debit": row[3],
            "amount": row[5] or row[4]
        })

    return row_data

def write_row_data(filepath, row_data):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    worksheet.title = "out"

    for row_raw_data in row_data:
        price = row_raw_data["amount"]

        row_1 = (row_raw_data["credit"], row_raw_data["name"], price)
        row_2 = (row_raw_data["debit"], row_raw_data["name"], price * -1)

        worksheet.append(row_1)
        worksheet.append(row_2)

    new_filename = os.path.splitext(filepath)[0] + "_OUTPUT.xlsx"

    workbook.save(new_filename)

def main(filepath):
    row_data = []

    if is_revolut(filepath):
        row_data = load_revolut(filepath)
    else:
        row_data = read_workbook(filepath)

    if not row_data:
        kill(1)
        return
    
    write_row_data(filepath, row_data)

def file_selection_callback(sender, appdata):
    main(appdata["file_path_name"])
    kill(0)

with imgui.file_dialog(directory_selector=False, show=False, tag="open_file_dialog", width=700, height=500, callback=file_selection_callback):
    imgui.add_file_extension(".xlsx")
    imgui.add_file_extension(".csv")

with imgui.window(label="Converter", width=800, height=600, no_title_bar=True):
    imgui.add_button(label="Valitse tiedosto...", callback=lambda: imgui.show_item("open_file_dialog"))
    imgui.add_input_int(label="Revolut credit account", tag="__revolut_credit_account", default_value=2880)
    imgui.add_input_int(label="Revolut topup account", tag="__revolut_topup_account", default_value=1700)
    imgui.add_input_int(label="Revolut debit account", tag="__revolut_debit_account", default_value=1930)

imgui.create_viewport(title="Converter", width=800, height=600)
imgui.setup_dearpygui()
imgui.show_viewport()
imgui.start_dearpygui()
imgui.destroy_context()