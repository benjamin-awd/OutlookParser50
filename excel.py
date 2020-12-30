import sqlite3
import xlsxwriter
import win32com.client as win32

import sys
import os

def main():
    workbook = xlsxwriter.Workbook("out.xlsx")
    worksheet = workbook.add_worksheet()

    path = os.getcwd()
    print(path)

    # Connect to database
    db = sqlite3.connect("emails.db")
    rows = db.execute("SELECT date, title, publication FROM articles")
    rows2 = db.execute("SELECT tier, category FROM articles")
    links = db.execute("SELECT * FROM links")
    platforms = db.execute("SELECT * FROM platforms")

    # Init formatting
    highlight, center_align, red_fill, bold = format(workbook)

    # Print header
    header = ["Date", "Title", "Publication", "In Tier 1?", "In Tier 2?", "In Tier 3?",
              "Business", "National", "Channel", "Trade", "Vertical", "Lifestyle"]
    worksheet.write_row("A1", header, bold)

    # Create variables to keep track of rows in excel and 'original' rows before extra rows are added
    row_counter = 0
    og_rows = []

    # Iterate through each row
    for db_row_num, row in enumerate(rows):
        print(f"\ndb_row_num is {db_row_num}")
        print(f"Current row on sheet = {row_counter}")

        extra_rows, link_list, platform_list = get_extra_rows(db_row_num, links, platforms)
        og_rows.append(row_counter)
        print(platform_list)

        # Iterate through items (e.g. date, title, publication) in row
        for item_num, item in enumerate(row):
            print(item_num, item)
            worksheet.write(row_counter + 1, item_num, item)

            # Special handling for title - hyperlinks
            if item_num == 1:
                title = item
                worksheet.write_url(row_counter + 1, item_num, link_list[0], string=f"{title}")

            # Special handling for publication - (publication type)
            if item_num == 2:
                publication = item
                worksheet.write(row_counter + 1, item_num, f"{publication} ({platform_list[0]})")

                for i in range(extra_rows):
                    # To write to current row below original row, add 1 to counter
                    row_counter += 1

                    # Iterate through items (e.g. date, title, publication) in row
                    for item_num, item in enumerate(row):
                        worksheet.write(row_counter + 1, item_num, item)

                        # Special handling for title - secondary hyperlinks (i + 1 as [0] already covered )
                        if item_num == 1:
                            title = item

                            try:
                                worksheet.write_url(row_counter + 1, item_num, link_list[i + 1], string=f"{title}")

                            # If out of range, don't hyperlink
                            except IndexError:
                                worksheet.write(row_counter + 1, item_num, "")

                        # Special handling for publication - secondary publication platforms (i + 1 as [0] already covered)
                        if item_num == 2:
                            publication = item

                            try:
                                worksheet.write(row_counter + 1, item_num, f"{publication} ({platform_list[i + 1]})")
                                print(f"  Type{i} = {platform_list[i + 1]}")

                            # If out of range, don't include additional type
                            except IndexError:
                                worksheet.write(row_counter + 1, item_num, f"{publication} ({platform_list[0]})")
                                print(f"   Type = {platform_list[0]}")

        # After iterating thorugh each original row, add 1 to row counter
        row_counter += 1

    for i, row in enumerate(rows2):
        tier = row[0]
        category = row[1]

        if tier == 1:
            worksheet.write(f"D{2 + og_rows[i]}", 1, center_align)

        if tier == 2:
            worksheet.write(f"E{2 + og_rows[i]}", 1, center_align)

        if tier == 3:
            worksheet.write(f"F{2 + og_rows[i]}", 1, center_align)

        if category == "Business":
            worksheet.write(f"G{2 + og_rows[i]}", 1, center_align)

        if category == "National":
            worksheet.write(f"H{2 + og_rows[i]}", 1, center_align)

        if category == "Channel":
            worksheet.write(f"I{2 + og_rows[i]}", 1, center_align)

        if category == "Trade":
            worksheet.write(f"J{2 + og_rows[i]}", 1, center_align)

        if category == "Vertical":
            worksheet.write(f"K{2 + og_rows[i]}", 1, center_align)

        if category == "Lifestyle":
            worksheet.write(f"L{2 + og_rows[i]}", 1, center_align)

    print(f"Last excel row = {row_counter}")
    worksheet.conditional_format(f"A1:C{row_counter}", {'type': 'blanks', 'format': highlight})
    worksheet.conditional_format(f"A1:C{row_counter}", {'type': 'text',
                                                        'criteria': 'containing', 'value': 'N/A', 'format': highlight})

    workbook.close()
    autofit(path)
    
def get_extra_rows(db_row_num, links, platforms):

    link_counter, link_list, link_id = get_links(links)
    type_counter, platform_list, article_id = get_platforms(platforms)

    if link_counter != type_counter:

        print(f"\nException detected: {link_counter} links, {type_counter} platforms in row {db_row_num + 1}")
        print(f"Links: {link_list}")
        print(f"Types: {platform_list}")

    if link_counter == type_counter:
        rows_to_print = link_counter

    if link_counter > type_counter:
        rows_to_print = link_counter

    if link_counter < type_counter:
        rows_to_print = type_counter

    return rows_to_print - 1, link_list, platform_list


def get_links(links):
    for row in links:
        link_counter = 0
        link_id = row[0]
        link_list = []

        for link in row:
            # Only capture links and not id
            if isinstance(link, str) == True:
                link_counter += 1

                # Capture each link (for debugging purposes)
                link_list.append(link)

            elif link == None:
                return link_counter, link_list, link_id

        print(row)


def get_platforms(platforms):
    for row in platforms:
        type_counter = 0
        type_id = row[0]
        platform_list = []

        for cat in row:
            # Only capture platforms and not id
            if isinstance(cat, str) == True:
                type_counter += 1

                # Capture each type (for debugging purposes)
                platform_list.append(cat)

            elif cat == None:
                return type_counter, platform_list, type_id

        print(row)


def autofit(path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(rf"{path}\out.xlsx")
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.Save()
    excel.Workbooks.Open(rf"{path}\out.xlsx")

    # Bring excel sheet to front and maximize (Source: https://stackoverflow.com/questions/19118881)
    excel.WindowState = win32.constants.xlMinimized
    excel.WindowState = win32.constants.xlMaximized


def format(workbook):
    bold = workbook.add_format()
    bold.set_bold()

    red_fill = workbook.add_format()
    red_fill.set_bg_color('red')

    center_align = workbook.add_format()
    center_align.set_align('center')

    highlight = workbook.add_format({'bg_color': '#FFFF00'})
    return highlight, center_align, red_fill, bold


main()