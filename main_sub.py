import requests
import json
# from openpyxl import Workbook, workbook
import openpyxl

from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile 
import time

ws = Tk()
ws.title('PythonGuides')
ws.geometry('400x200') 

def getApi():
    chp = requests.get("https://script.googleusercontent.com/macros/echo?user_content_key=Pc6Nn6aImypsxrVm-1z70insSpR0JbWuRhkrV5Mua08oi1NucVZ8Bkl1UfGF3nGK5Hiu7vhCa23ZlElY8R-RrE29JUEmNOcUm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnM4eD-czAlqGo1zibcjBOW-XtK_X5NhfCKWFBlwkZ76HhUymyjaR76ButyWRpXn7fbpz3ZducOMor69vOtJkWh23MRqSkdZDtdz9Jw9Md8uu&lib=M7Gjc0xFmBFKHNMTgDUmVi1-DpKjjLho_")
    print(chp.text)

    rng = requests.get("https://script.google.com/macros/s/AKfycbzVCiHUba5ymw7kHH7A2b6wJYshXNCHEGMKJpcCizWYGjQmDDokxkFX09lMThnkhaLhow/exec")
    print(rng.text)

    pkk = requests.get("https://script.google.com/macros/s/AKfycbyPhMNTrQyHT8pWImJIkneIAqwNp4S9rs5KfPp1ihnrJTqXHgmlDCm3t9HNnbbs9_wZMA/exec")
    print(pkk.text)

    hua = requests.get("https://script.google.com/macros/s/AKfycbwo89Y_iCUpqambco8XuHKadQFftCw2KQDP12FxhnNsx3Cxjy-buCYLcekxeC5Gq4bnpw/exec")
    print(hua.text)

    pnp = requests.get("https://script.google.com/macros/s/AKfycbxkGn6DudZod0bC_7gn7743EyQPs8vCK4QowFq-2hyQwaSBAt9Hlwls2gODlzIhymrm/exec")
    print(pnp.text)

    pbi = requests.get("https://script.google.com/macros/s/AKfycby0Y-TCraFEPav_oNkzlN18jBO0nyxD1tf9S8d-H3oG7_Kap90QgCnuWSr5SnDkghuALg/exec")
    print(pbi.text)

    tsk = requests.get("https://script.google.com/macros/s/AKfycbwfxqGahrP6AUiDOgthi_IFI7c2OYRQAp08MLyh7p66BKwYtPQDd7JIjqDL_CNi6Zcq/exec")
    print(tsk.text)


    lhs = requests.get("https://script.google.com/macros/s/AKfycbygeimfLYN2AjKlGqlcK34IiP3KUB8hxA-jnS4QNzwybCV3dlp7xYMi3xJ5cLKYhqSxMw/exec")
    print(lhs.text)

    swi = requests.get("https://script.google.com/macros/s/AKfycbx21vkXLpjGKxsuFja2R91FCYMHdHe5mnPt6De9kLdzBUWFwyz576MFQn7tLgwEx1Pg0g/exec")
    print(swi.text)

    kbr = requests.get("https://script.google.com/macros/s/AKfycbzKoUAk5PLsS4hJE1W80OHG_5MF8OPjjwr4jFtj18Q8LTsRiUbyxZfN2SM8VW-AJDYOSA/exec")
    print(kbr.text)

    kpo = requests.get("https://script.google.com/macros/s/AKfycbxJp43n4ncEZdE8kIhOYFiG1L3eTpr1aySs1RFjoGsMG1zhsoUJ4jIXkFgRVemWahNmfw/exec")
    print(kpo.text)

    pht = requests.get("https://script.google.com/macros/s/AKfycbzMfKq_QK0XeDGkC4vUs0cUKTahg5sOYhLwufKERKtXtJHKA4WVG54CHPFJXHhYrM6qBg/exec")
    print(pht.text)

    kura = requests.get("https://script.google.com/macros/s/AKfycbwfKZWykbuBtKKnFBnGLA2_e7H-4mnSZo7t87FhR_EpFyafDyZYUFq9PtjhlKjdtQPQ/exec")
    print(kura.text)

    tas = requests.get("https://script.google.com/macros/s/AKfycby8lWjdTLW7PxIK3oR8PbhL-WH1A1QSqFl7cQTiNpSCtDnvrw9aA5Z-0NhdO_fLHqsh/exec")
    print(tas.text)
    

    json_data_chp = json.loads(chp.text) # print(json_data)
    json_data_rng = json.loads(rng.text) 
    json_data_pkk = json.loads(pkk.text) 
    json_data_hua = json.loads(hua.text) 
    json_data_pnp = json.loads(pnp.text) 
    json_data_pbi = json.loads(pbi.text) 
    json_data_lhs = json.loads(lhs.text)
    json_data_swi = json.loads(swi.text)
    json_data_kbr = json.loads(kbr.text)
    json_data_kpo = json.loads(kpo.text)
    json_data_pht = json.loads(pht.text)
    json_data_kura = json.loads(kura.text)
    json_data_tas = json.loads(tas.text)
    json_data_tsk = json.loads(tsk.text)


    upDateTime = json_data_chp[-1]['DateTime']

    # #POWER Mux OR A Watt
    # CHP_PowerW = json_data_chp[-1]['PowerMux(W)']  
    # RNG_PowerW = json_data_rng[-1]['PowerMux(W)'] 
    # PKK_PowerW = json_data_pkk[-1]['PowerMux(W)'] 
    # HUA_PowerW = json_data_hua[-1]['PowerMux(W)'] 
    # PNP_PowerW = json_data_pnp[-1]['PowerMux(W)']
    # PBI_PowerW = json_data_pbi[-1]['PowerA(W)']
    # LHS_PowerW = json_data_lhs[-1]['PowerA(W)']
    # SWI_PowerW = json_data_swi[-1]['PowerA(W)']    
    # KBR_PowerW = json_data_kbr[-1]['Power(W)']
    # KPO_PowerW = json_data_kpo[-1]['Power(W)']
    # PHT_PowerW = json_data_pht[-1]['Power(W)']
    # KURA_PowerW = json_data_kura[-1]['Power(W)']
    # TAS_PowerW = json_data_tas[-1]['PowerA(W)']

    # #POWER B Watt
    # PBI_PowerBW = json_data_pbi[-1]['PowerB(W)']
    # LHS_PowerBW = json_data_lhs[-1]['PowerB(W)']
    # TAS_PowerBW = json_data_tas[-1]['PowerB(W)']

    #IRD
    CHP_IRD_CN_A = json_data_chp[-1]['IRD_A_C/N'] 
    CHP_IRD_LM_A = json_data_chp[-1]['IRD_A_LM']
    CHP_IRD_CN_B = json_data_chp[-1]['IRD_B_C/N'] 
    CHP_IRD_LM_B = json_data_chp[-1]['IRD_B_LM']

    RNG_IRD_CN_A = json_data_rng[-1]['IRD_A_C/N'] 
    RNG_IRD_LM_A = json_data_rng[-1]['IRD_A_LM']
    RNG_IRD_CN_B = json_data_rng[-1]['IRD_B_C/N'] 
    RNG_IRD_LM_B = json_data_rng[-1]['IRD_B_LM']

    PKK_IRD_CN_A = json_data_pkk[-1]['IRD_A_C/N'] 
    PKK_IRD_LM_A = json_data_pkk[-1]['IRD_A_LM'] 
    PKK_IRD_CN_B = json_data_pkk[-1]['IRD_B_C/N'] 
    PKK_IRD_LM_B = json_data_pkk[-1]['IRD_B_LM']

    HUA_IRD_CN_A = json_data_hua[-1]['IRD_A_C/N'] 
    HUA_IRD_LM_A = json_data_hua[-1]['IRD_A_LM']
    HUA_IRD_CN_B = json_data_hua[-1]['IRD_B_C/N'] 
    HUA_IRD_LM_B = json_data_hua[-1]['IRD_B_LM']

    TSK_IRD_CN_A = json_data_tsk[-1]['IRD_A_CN'] 
    TSK_IRD_LM_A = json_data_tsk[-1]['IRD_A_LM']
    TSK_IRD_CN_B = json_data_tsk[-1]['IRD_B_CN'] 
    TSK_IRD_LM_B = json_data_tsk[-1]['IRD_B_LM']

    PNP_IRD_CN = json_data_pnp[-1]['IRD_A_CN'] 
    PNP_IRD_LM = json_data_pnp[-1]['IRD_A_LM']

    PBI_IRD_CN_A = json_data_pbi[-1]['IRD_A_CN'] 
    PBI_IRD_LM_A = json_data_pbi[-1]['IRD_A_LM'] 
    PBI_IRD_CN_B = json_data_pbi[-1]['IRD_B_CN'] 
    PBI_IRD_LM_B = json_data_pbi[-1]['IRD_B_LM'] 

    LHS_IRD_CN_A = json_data_lhs[-1]['IRD_A_CN'] 
    LHS_IRD_LM_A = json_data_lhs[-1]['IRD_A_LM']
    LHS_IRD_CN_B = json_data_lhs[-1]['IRD_B_CN'] 
    LHS_IRD_LM_B = json_data_lhs[-1]['IRD_B_LM'] 

    SWI_IRD_CN = json_data_swi[-1]['IRD_A_CN'] 
    SWI_IRD_LM = json_data_swi[-1]['IRD_A_LM']  

    KBR_IRD_CN = json_data_kbr[-1]['IRD_A_CN'] 
    KBR_IRD_LM = json_data_kbr[-1]['IRD_A_LM']

    KPO_IRD_CN = json_data_kpo[-1]['IRD_A_CN'] 
    KPO_IRD_LM = json_data_kpo[-1]['IRD_A_LM']

    PHT_IRD_CN = json_data_pht[-1]['IRD_A_CN'] 
    PHT_IRD_LM = json_data_pht[-1]['IRD_A_LM']  

    KURA_IRD_CN = json_data_kura[-1]['IRD_A_CN'] 
    KURA_IRD_LM = json_data_kura[-1]['IRD_A_LM']  

    TAS_IRD_CN_A = json_data_tas[-1]['IRD_A_CN'] 
    TAS_IRD_LM_A = json_data_tas[-1]['IRD_A_LM'] 
    TAS_IRD_CN_B = json_data_tas[-1]['IRD_B_CN'] 
    TAS_IRD_LM_B = json_data_tas[-1]['IRD_B_LM']  
        

    # #UPS
    CHP_UpsRuntime = json_data_chp[-1]['UPSRuntime']
    RNG_UpsRuntime = json_data_rng[-1]['UPSRuntime']
    PKK_UpsRuntime = json_data_pkk[-1]['UPSRuntime']
    TSK_UpsRuntime = json_data_tsk[-1]['UPSRuntime']
    PHT_UpsRuntime = json_data_pht[-1]['UPS_Runtime']
    KPO_UpsRuntime = json_data_kpo[-1]['UPS_Runtime']
    KURA_UpsRuntime = json_data_kura[-1]['UPS_Runtime']


    print(upDateTime)
    # print('ชุมพร', ' : ', CHP_PowerW, 'W')
    # print('ระนอง', ' : ', RNG_PowerW, 'W')
    # print('ประจวบ', ' : ', PKK_PowerW, 'W')
    # print('หัวหิน', ' : ', HUA_PowerW, 'W')
    # print('ปากน้ำปราณ', ' : ', PNP_PowerW, 'W')
    # print('เพชรบุรี' ,' : ', PBI_PowerW , 'W')
    # print('หลังสวน', ' : ', LHS_PowerW , 'W')
    # print('สวี', ' : ', SWI_PowerW , 'W')
    # print('กระบุรี' ' : ', KBR_PowerW , 'W')
    # print('กะเปอร์' ' : ', KPO_PowerW , 'W')
    # print('พะโต๊ะ' ' : ', PHT_PowerW , 'W')
    # print('คุระบุรี' ' : ', KURA_PowerW , 'W')
    # print('ท่าแซะ' ' : ', TAS_PowerW , 'W')

    #IRD
    ird_chp_cn_a = str(CHP_IRD_CN_A)
    ird_chp_lm_a = str(CHP_IRD_LM_A)
    ird_chp_cn_b = str(CHP_IRD_CN_B)
    ird_chp_lm_b = str(CHP_IRD_LM_B)

    ird_rng_cn_a = str(RNG_IRD_CN_A) 
    ird_rng_lm_a = str(RNG_IRD_LM_A) 
    ird_rng_cn_b = str(RNG_IRD_CN_B) 
    ird_rng_lm_b = str(RNG_IRD_LM_B) 

    ird_pkk_cn_a = str(PKK_IRD_CN_A) 
    ird_pkk_lm_a = str(PKK_IRD_LM_A)
    ird_pkk_cn_b = str(PKK_IRD_CN_B) 
    ird_pkk_lm_b = str(PKK_IRD_LM_B)

    ird_hua_cn_a = str(HUA_IRD_CN_A)
    ird_hua_lm_a = str(HUA_IRD_LM_A)
    ird_hua_cn_b = str(HUA_IRD_CN_B)
    ird_hua_lm_b = str(HUA_IRD_LM_B)

    ird_tsk_cn_a = str(TSK_IRD_CN_A)
    ird_tsk_lm_a = str(TSK_IRD_LM_A)
    ird_tsk_cn_b = str(TSK_IRD_CN_B)
    ird_tsk_lm_b = str(TSK_IRD_LM_B)


    ird_pnp_cn = str(PNP_IRD_CN)
    ird_pnp_lm = str(PNP_IRD_LM)

    ird_pbi_cn_a = str(PBI_IRD_CN_A) 
    ird_pbi_lm_a = str(PBI_IRD_LM_A) 
    ird_pbi_cn_b = str(PBI_IRD_CN_B) 
    ird_pbi_lm_b = str(PBI_IRD_LM_B) 

    ird_lhs_cn_a = str(LHS_IRD_CN_A)
    ird_lhs_lm_a = str(LHS_IRD_LM_A)
    ird_lhs_cn_b = str(LHS_IRD_CN_B)
    ird_lhs_lm_b = str(LHS_IRD_LM_B)

    ird_swi_cn = str(SWI_IRD_CN) 
    ird_swi_lm = str(SWI_IRD_LM)

    ird_kbr_cn = str(KBR_IRD_CN) 
    ird_kbr_lm = str(KBR_IRD_LM)

    ird_kpo_cn = str(KPO_IRD_CN) 
    ird_kpo_lm = str(KPO_IRD_LM) 

    ird_pht_cn = str(PHT_IRD_CN)
    ird_pht_lm = str(PHT_IRD_LM)

    ird_kura_cn = str(KURA_IRD_CN) 
    ird_kura_lm = str(KURA_IRD_LM)

    ird_tas_cn_a = str(TAS_IRD_CN_A) 
    ird_tas_lm_a = str(TAS_IRD_LM_A) 
    ird_tas_cn_b = str(TAS_IRD_CN_B) 
    ird_tas_lm_b = str(TAS_IRD_LM_B) 

    # #UPS
    ups_runtime_chp = str(CHP_UpsRuntime) 
    ups_runtime_rng = str(RNG_UpsRuntime)
    ups_runtime_pkk = str(PKK_UpsRuntime)
    ups_runtime_pht = str(PHT_UpsRuntime)
    ups_runtime_kpo = str(KPO_UpsRuntime)
    ups_runtime_kura = str(KURA_UpsRuntime)
    ups_runtime_tsk = str(TSK_UpsRuntime)


    # filename = "report.xlsx"

    #workbook = Workbook()
    # sheet = workbook.active

    # #CHP ird 
    # sheet["AA33"] = ird_chp_lm_a
    # sheet["AB33"] = ird_chp_cn_a
    # sheet["AC33"] = ird_chp_lm_b
    # sheet["AD33"] = ird_chp_cn_b

    # workbook.save(filename=filename)


    xfile = openpyxl.load_workbook('report.xlsx')

    sheet = xfile.get_sheet_by_name('Daily report')
    sheet["AA33"] = ird_chp_lm_a
    sheet["AB33"] = ird_chp_cn_a
    sheet["AC33"] = ird_chp_lm_b
    sheet["AD33"] = ird_chp_cn_b
    sheet["Y33"] = ups_runtime_chp

    sheet["AA34"] = ird_rng_lm_a
    sheet["AB34"] = ird_rng_cn_a
    sheet["AC34"] = ird_rng_lm_b
    sheet["AD34"] = ird_rng_cn_b
    sheet["Y34"] = ups_runtime_rng

    sheet["AA35"] = ird_pkk_lm_a
    sheet["AB35"] = ird_pkk_cn_a
    sheet["AC35"] = ird_pkk_lm_b
    sheet["AD35"] = ird_pkk_cn_b
    sheet["Y35"] = ups_runtime_pkk

    sheet["AA36"] = ird_tas_lm_a
    sheet["AB36"] = ird_tas_cn_a
    sheet["AC36"] = ird_tas_lm_b
    sheet["AD36"] = ird_tas_cn_b

    sheet["AA37"] = ird_lhs_lm_a
    sheet["AB37"] = ird_lhs_cn_a
    sheet["AC37"] = ird_lhs_lm_b
    sheet["AD37"] = ird_lhs_cn_b

    sheet["AA38"] = ird_tsk_lm_a
    sheet["AB38"] = ird_tsk_cn_a
    sheet["AC38"] = ird_tsk_lm_b
    sheet["AD38"] = ird_tsk_cn_b
    sheet["Y38"] = ups_runtime_tsk

    sheet["AA39"] = ird_pbi_lm_a
    sheet["AB39"] = ird_pbi_cn_a
    sheet["AC39"] = ird_pbi_lm_b
    sheet["AD39"] = ird_pbi_cn_b

    sheet["AA40"] = ird_hua_lm_a
    sheet["AB40"] = ird_hua_cn_a
    sheet["AC40"] = ird_hua_lm_b
    sheet["AD40"] = ird_hua_cn_b

    sheet["AA41"] = ird_pht_lm
    sheet["AB41"] = ird_pht_cn
    sheet["Y41"] = ups_runtime_pht

    sheet["AA42"] = ird_kpo_lm
    sheet["AB42"] = ird_kpo_cn
    sheet["Y42"] = ups_runtime_kpo

    sheet["AA43"] = ird_kura_lm
    sheet["AB43"] = ird_kura_cn
    sheet["Y43"] = ups_runtime_kura

    sheet["AA44"] = ird_swi_lm
    sheet["AB44"] = ird_swi_cn

    sheet["AA45"] = ird_pnp_lm
    sheet["AB45"] = ird_pnp_cn

    sheet["AA46"] = ird_kbr_lm
    sheet["AB46"] = ird_kbr_cn

    xfile.save('report2.xlsx')


def uploadFiles():
    pb1 = Progressbar(
        ws, 
        orient=HORIZONTAL, 
        length=300, 
        mode='determinate'
        )
    pb1.grid(row=4, columnspan=3, pady=20)
    for i in range(10):
        ws.update_idletasks()
        pb1['value'] += 10
        time.sleep(1)
    pb1.destroy()
    Label(ws, text='File Uploaded Successfully!', foreground='green').grid(row=4, columnspan=3, pady=10)


upld = Button(
ws, 
text='Upload Files', 
command = lambda:[uploadFiles(),getApi()])

upld.grid(row=2, columnspan=3, pady=5)

ws.mainloop()