import tkinter as tk
from tkinter import filedialog, ttk
from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import os
import pandas as pd
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from itertools import chain
from numpy import nan

st = 0
driver = None

def start(month="All", year="All", srow=0, override=0):
    if outfile == "" or file == "":
        tk.messagebox.showerror(message='Please select input file and output folder!\nGo to Settings tab.')
        return
    
    global st
    st = 0
    fname = outfile + "/GST" + time.strftime("%d%m%Y-%H%M") + ".csv"
    entries = pd.read_excel(file)
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
              "November", "December"]

    global driver
    driver = webdriver.Chrome()
    driver.get("https://services.gst.gov.in/services/searchtp")
    driver.maximize_window()
    if month == "All" and year == "All":
        f = open(fname, "a")
        f.write(",GSTIN,NAME,January,February,March,April,May,June,July,August,September,October,November,December\n")
        f.close()
        for i, j, p in zip(entries["GSTIN"], entries["NAMES"], entries["STATUS"]):

            if p == "Y" and override == 0:
                continue

            if srow > 2:
                srow -= 1
                continue

            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "for_gstin")))
            finally:
                idinput = driver.find_element(By.ID, "for_gstin")
                idinput.send_keys(i)

            if st == 1:
                break

            a = 0
            while (a == 0):
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "filingTable")))
                except:
                    pass
                else:
                    a = 1
                    tablebut = driver.find_element(By.ID, "filingTable")
                    tablebut.click()

            while (len(driver.find_elements(By.CLASS_NAME, "table")) <= 1):
                time.sleep(0.1)
            table = driver.find_elements(By.CLASS_NAME, "table")

            t1 = table[1].text
            t2 = table[2].text
            t1 = [x.split(",") for x in t1.replace(" ", ",")[47:].split("\n")]
            t2 = [x.split(",") for x in t2.replace(" ", ",")[47:].split("\n")]
            t1 = list(chain.from_iterable(t1))
            t2 = list(chain.from_iterable(t2))

            df = pd.DataFrame(columns=['GSTIN', 'NAME', 'January', 'February', 'March', 'April',
                              'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'])
            flag = [0]*12
            tem = 0
            for k in range(len(t1)):
                if t1[k] in months:
                    tem = months.index(t1[k])
                    flag[tem] += 1
                    df.loc[2*flag[tem] - 1, t1[k]] = t1[k-1]
                    df.loc[2*flag[tem], t1[k]] = t1[k+1]

            mflag = max(flag)
            flag = [0]*12
            df.loc[2*mflag + 1, :] = nan
            for k in range(len(t2)):
                if t2[k] in months:
                    tem = months.index(t2[k])
                    flag[tem] += 1
                    df.loc[2*mflag + 2*flag[tem], t2[k]] = t2[k-1]
                    df.loc[2*mflag + 2*flag[tem] + 1, t2[k]] = t2[k+1]
            df.loc[2*mflag + 2*flag[tem] + 2, :] = nan
            df.loc[1, 'GSTIN'] = i
            df.loc[1, 'NAME'] = j
            df.to_csv(fname, mode='a', header=False)
            entries.loc[entries["GSTIN"] == i, "STATUS"] = "Y"

            del df
        df = pd.read_csv(fname)
        os.remove(fname)
        df.to_excel(str(fname)[:-3]+"xlsx", index=False)
        del df
        entries.to_excel(file, index=False)

    else:
        tem = 0
        df = pd.DataFrame(
            columns=['GSTIN', 'NAME', month+" " +year+" 3B", month+" "+year+" 1/IFF"])
        for i, j, p in zip(entries["GSTIN"], entries["NAMES"], entries["STATUS"]):
            if p == "Y" and override == 0:
                continue

            if srow > 2:
                srow -= 1
                continue

            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "for_gstin")))
            finally:
                idinput = driver.find_element(By.ID, "for_gstin")
                idinput.send_keys(i)
            
            if st == 1:
                break
            
            a = 0
            while (a == 0):
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, "filingTable")))
                except:
                    pass
                else:
                    a = 1
                    tablebut = driver.find_element(By.ID, "filingTable")
                    tablebut.click()

            while (len(driver.find_elements(By.CLASS_NAME, "table")) <= 1):
                time.sleep(0.1)
            table = driver.find_elements(By.CLASS_NAME, "table")

            t1 = table[1].text
            t2 = table[2].text
            t1 = [x.split(",") for x in t1.replace(" ", ",")[47:].split("\n")]
            t2 = [x.split(",") for x in t2.replace(" ", ",")[47:].split("\n")]
            t1 = list(chain.from_iterable(t1))
            t2 = list(chain.from_iterable(t2))
            for k in range(len(t1)):
                if t1[k] == month and t1[k-1] == year:
                    df.loc[tem, month+" " +year+" 3B"] = t1[k+1]

            for k in range(len(t2)):
                if t2[k] == month and t2[k-1] == year:
                    df.loc[tem, month+" " +year+" 1/IFF"] = t2[k+1]

            df.loc[tem, 'GSTIN'] = i
            df.loc[tem, 'NAME'] = j

            tem += 1
            entries.loc[entries["GSTIN"] == i, "STATUS"] = "Y"
        df.to_excel(str(fname)[:-3]+"xlsx", index=False)
        del df
        entries.to_excel(file, index=False)

    tk.messagebox.showinfo(message='Completed!')
    driver.quit()
    del driver
    os.startfile(str(fname)[:-3]+"xlsx")
    root.quit()
    
    return


def outFolderButton():
    global outfile
    outfile = filedialog.askdirectory()
    filenameLabel["text"] = "Output Folder: {}".format(outfile)
    con = open("config.txt", "w")
    con.write(str(outfile)+"\n")
    con.write(str(file)+"\n")
    con.close()
    return


def inpFileButton():
    global file
    file = filedialog.askopenfilename()
    inpfileLabel["text"] = "Input File: {}".format(file)
    con = open("config.txt", "w")
    con.write(str(outfile)+"\n")
    con.write(str(file)+"\n")
    con.close()
    return


def thr(month, year, srow, override):
    t = Thread(target=start, args=(month, year, srow, override))
    t.start()
    return


def stop():
    # askyes no
    ans = tk.messagebox.askyesno(title="Abort", message="Are you sure you want to abort?")
    if ans == True:
        global st
        st = 1
        global driver
        driver.quit()
        return
    return

def createSampleFile():
    if outfile == "":
        tk.messagebox.showerror(message='Please select output folder!\nGo to Settings tab.')
        return
    fname = outfile + "/Sample Input File.xlsx"
    df = pd.DataFrame(columns=['GSTIN', 'NAMES', 'STATUS'])
    try:
        df.to_excel(fname, index=False)
    except:
        tk.messagebox.showerror(message='File already exists!')
    tk.messagebox.showinfo(message='Sample Input File created!')
    return


if __name__ == "__main__":
    root = tk.Tk()
    root.title("GST")

    tabControl = ttk.Notebook(root)
    run = ttk.Frame(tabControl)
    settings = ttk.Frame(tabControl)
    about = ttk.Frame(tabControl)
    tabControl.add(run, text='Run')
    tabControl.add(settings, text='Settings')
    tabControl.add(about, text='About')
    tabControl.grid(row=0, column=0, padx=50, pady=10)

    # load config
    try:
        con = open("config.txt", "r")
        outfile = con.readline()[:-1]
        file = con.readline()[:-1]
        con.close()
    except:
        outfile = ""
        file = ""

    # output folder
    filenameLabel = tk.Label(
        settings, text="Output Folder: {}".format(outfile), justify="left")
    filenameLabel.grid(row=0, column=0, padx=5, pady=10)
    outfileButton = tk.Button(settings, text="Select",
                              width=15, command=lambda: outFolderButton())
    outfileButton.grid(row=0, column=1, padx=5, pady=10)

    # input file
    inpfileLabel = tk.Label(settings, text="Input File: {}".format(file), justify="left")
    inpfileLabel.grid(row=1, column=0, padx=5, pady=10)
    inpfileButton = tk.Button(settings, text="Select",
                              width=15, command=lambda: inpFileButton())
    inpfileButton.grid(row=1, column=1, padx=5, pady=10)

    # create sample input file button
    sampleButton = tk.Button(settings, text="Create Sample Input File", width=20, command=lambda: createSampleFile())
    sampleButton.grid(row=2, column=0, padx=0, pady=10, )

    # dropdown for specific month and year
    monthLabel = tk.Label(run, text="Month:")
    monthLabel.grid(row=1, column=0, padx=5, pady=10)
    month = tk.StringVar(root)
    month.set("All")
    monthDropdown = tk.OptionMenu(run, month, "All", "January", "February", "March", "April",
                                  "May", "June", "July", "August", "September", "October", "November", "December")
    monthDropdown.grid(row=1, column=1, padx=5, pady=10)

    yearLabel = tk.Label(run, text="Year:")
    yearLabel.grid(row=2, column=0, padx=5, pady=10)
    year = tk.StringVar(root)
    year.set("All")
    yearDropdown = tk.OptionMenu(run, year, "All", "2017-2018", "2018-2019",
                                 "2019-2020", "2020-2021", "2021-2022", "2022-2023", "2023-2024")
    yearDropdown.grid(row=2, column=1, padx=5, pady=10)

    # startrow entry
    startrowLabel = tk.Label(run, text="Start Row:")
    startrowLabel.grid(row=3, column=0, padx=5, pady=10)
    startrowEntry = tk.Entry(run, width=15)
    startrowEntry.insert(0, "2")
    startrowEntry.grid(row=3, column=1, padx=5, pady=10)

    # override status chechbox
    override = tk.IntVar()
    overrideCheck = tk.Checkbutton(run, text="Override Status", variable=override, onvalue=1, offvalue=0)
    overrideCheck.grid(row=4, column=0, padx=5, pady=10)

    # start button
    startButton = tk.Button(run, text="Start", width=15,
                            command=lambda: thr(month.get(), year.get(), int(startrowEntry.get()), override.get()))
    startButton.grid(row=5, column=0, padx=5, pady=10)

    # abort button
    abortButton = tk.Button(run, text="Abort", width=15,
                            command=lambda: stop())
    abortButton.grid(row=5, column=1, padx=5, pady=10)

    # about
    aboutLabel = tk.Label(about, text="Developed by: Prasanna Paithankar\nfor Bhushan & Associates\nVersion: 1.0.0 (2023)", justify="left")
    aboutLabel.grid(row=0, column=0, padx=5, pady=10)

    root.mainloop()
