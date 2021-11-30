from tkinter import *
from tkinter.ttk import *
from tkinter import *
from tkinter.colorchooser import *
from tkinter.filedialog import *
from tkinter.ttk import Progressbar
import tkinter as tk
from tkinter import ttk
import requests
import json
import openpyxl
import os.path
import time
import shutil
import threading
import pygame
import datetime


pygame.mixer.init()
def play():
	# playsound('sound1.mp3')
	pygame.mixer.music.load("search.mp3")
	pygame.mixer.music.play(loops=0)

def play2():
    	# playsound('sound1.mp3')
	pygame.mixer.music.load("sound1.mp3")
	pygame.mixer.music.play(loops=5)     


def playThread():
	task = threading.Thread(target=play2)
	task.start()           


def selectFile():
        playThread()
        try:
                
                # fileopen = askopenfile()
                # myLabel = Label(text=fileopen).pack()
                # fileopen = askopenfilename()
                fileContent = askopenfilename()
                # fileContent = open(fileopen,encoding="UTF-8")
                save_path = 'E:/autoReport'
                name_of_file = fileContent
                completeName = os.path.join(save_path, name_of_file) 
                
                myLabel = Label(text=fileContent).pack()

                if (fileContent != ''):
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
                        
                        chp_swr = requests.get("https://script.google.com/macros/s/AKfycbxxCBgkp11AgaexWczzbLnp29Vryt73kc6Zmbz0v_ix8-1v51ZJzklUMltAYztLuQ4WQQ/exec")
                        print(chp_swr.text)

                        tsk_swr = requests.get("https://script.google.com/macros/s/AKfycbysVPXw6jFlvvsyFy-Mrr3vYJ7iBZH2B0B1BuPvX2suUC_AwxxuG0HjjQ2fjc3NzMnG/exec")
                        print(tsk_swr.text)

                        rng_swr = requests.get("https://script.google.com/macros/s/AKfycbw4OYLAEdlG5ND-YB5gCUH-V-ZPpxtFQhGaKXT_w_UcvtnAH2yrOV-vw1pt-bF5GWnY/exec")
                        print(rng_swr.text)

                        pht_swr = requests.get("https://script.google.com/macros/s/AKfycbwxIectpz3ghwaBrgcdLoJRm2uSGxYR0JkCnSAwbocue3_xLfJAnry2lKscBR3R8K4/exec")
                        print(pht_swr.text)

                        kpo_swr = requests.get("https://script.google.com/macros/s/AKfycbwffBxTTDMUMl-t32IiIYc6bJGq6o92VjynhMN2p3Gk_ReuB-hpiPbNb3KUms6gWKtK/exec")
                        print(kpo_swr.text)

                        kai_swr = requests.get("https://script.google.com/macros/s/AKfycbwPRvT9A9GKOHzW86YPuP2W5UF1YszI6gcar2pZa3fxdnK4LDOfxx9Mvsqt408SsU-t/exec")
                        print(kai_swr.text)

                        sfn = requests.get("https://script.google.com/macros/s/AKfycbybWq01yB97aDwjfVn7BUo1ne1nyEb_zfYw9OQihQBIL6LfZRP360yLwJGlQyvjG3BNiQ/exec")
                        print(sfn.text)

                        loadUps = requests.get("https://script.google.com/macros/s/AKfycbw-p73czCl28McEXEVpRohg7tfTLrIvpqRm8ZuInp0b918ybn5GicKMgZFHufP-T13IBw/exec")
                        print(loadUps.text)


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
                        json_data_chp_swr = json.loads(chp_swr.text)
                        json_data_tsk_swr = json.loads(tsk_swr.text)
                        json_data_rng_swr = json.loads(rng_swr.text)
                        json_data_pht_swr = json.loads(pht_swr.text)
                        json_data_kpo_swr = json.loads(kpo_swr.text)
                        json_data_kai_swr = json.loads(kai_swr.text)
                        json_data_sfn = json.loads(sfn.text)
                        json_data_loadUps = json.loads(loadUps.text)


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

                        #SFN
                        SFN_MARGIN_LHS = json_data_sfn[-1]['SFN_LHS']
                        SFN_MARGIN_TAS = json_data_sfn[-1]['SFN_TAS']
                        SFN_MARGIN_KPO = json_data_sfn[-1]['SFN_KPO']
                        SFN_MARGIN_KAI = json_data_sfn[-1]['SFN_KAI']
                        SFN_MARGIN_PNP = json_data_sfn[-1]['SFN_PNP']

                        #Emmission time
                        EM_LHS = json_data_sfn[-1]['EMMIS_LHS']
                        EM_TAS = json_data_sfn[-1]['EMMIS_TAS']
                        EM_KPO = json_data_sfn[-1]['EMMIS_KPO']
                        EM_KAI = json_data_sfn[-1]['EMMIS_KAI']
                        EM_PNP = json_data_sfn[-1]['EMMIS_PNP']

                        #SWR
                        CHP_UPPER   = json_data_chp_swr[-1]['upper']
                        CHP_LOWER   = json_data_chp_swr[-1]['lower']
                        CHP_FWLOWER = json_data_chp_swr[-1]['FWD_lower']
                        CHP_FWUPPER = json_data_chp_swr[-1]['FWD_upper']

                        TSK_UPPER = json_data_tsk_swr[-1]['upper']
                        TSK_LOWER   = json_data_tsk_swr[-1]['lower']
                        TSK_FWLOWER = json_data_tsk_swr[-1]['FWD_lower']
                        TSK_FWUPPER = json_data_tsk_swr[-1]['FWD_upper']

                        RNG_UPPER = json_data_rng_swr[-1]['upper']
                        RNG_LOWER   = json_data_rng_swr[-1]['lower']
                        RNG_FWLOWER = json_data_rng_swr[-1]['FWD_lower']
                        RNG_FWUPPER = json_data_rng_swr[-1]['FWD_upper']

                        PHT_FWANT = json_data_pht_swr[-1]['FWD_ant']
                        PHT_SWRANT = json_data_pht_swr[-1]['SWR_ant']

                        KPO_FWANT = json_data_kpo_swr[-1]['FWD_ant']
                        KPO_SWRANT = json_data_kpo_swr[-1]['SWR_ant']

                        KAI_FWANT = json_data_kai_swr[-1]['FWD_ant']
                        KAI_SWRANT = json_data_kai_swr[-1]['SWR_ant']


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

                        # load ups %
                        CHP_load_ups = json_data_loadUps[-1]['loadCHP']
                        RNG_load_ups = json_data_loadUps[-1]['loadRNG']
                        PKK_load_ups = json_data_loadUps[-1]['loadPKK']
                        TSK_load_ups = json_data_loadUps[-1]['loadTSK']

                



                        print(upDateTime)
                        # print(RNG_UpsRuntime)
                        
                        #Convert minute to hours rng station
                        rngTimehours = int(RNG_UpsRuntime / 60)
                        rngTimeMinutes = int(RNG_UpsRuntime - 60)
                        # print('Hours: ',rngTimehours , '\n', 
                        # 'Minutes: ', rngTimeMinutes)
         

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

                        #UPS
                        ups_runtime_chp = str(CHP_UpsRuntime) 
                        ups_runtime_rng_h = str(rngTimehours)
                        ups_runtime_rng_min = str(rngTimeMinutes)
                        ups_runtime_pkk = str(PKK_UpsRuntime)
                        ups_runtime_pht = str(PHT_UpsRuntime)
                        ups_runtime_kpo = str(KPO_UpsRuntime)
                        ups_runtime_kura = str(KURA_UpsRuntime)
                        ups_runtime_tsk = str(TSK_UpsRuntime)

                        #load ups % to string
                        chp_load_ups = str(CHP_load_ups)
                        rng_load_ups = str(RNG_load_ups)
                        pkk_load_ups = str(PKK_load_ups)
                        tsk_load_ups = str(TSK_load_ups)


                        #SWR to string
                        chp_upper = str(CHP_UPPER)
                        chp_lower = str(CHP_LOWER)
                        chp_fwupper = str(CHP_FWUPPER)
                        chp_fwlower = str(CHP_FWLOWER)

                        rng_upper = str(RNG_UPPER)
                        rng_lower = str(RNG_LOWER)
                        rng_fwupper = str(RNG_FWUPPER)
                        rng_fwlower = str(RNG_FWLOWER)

                        tsk_upper = str(TSK_UPPER)
                        tsk_lower = str(TSK_LOWER)
                        tsk_fwupper = str(TSK_FWUPPER)
                        tsk_fwlower = str(TSK_FWLOWER)

                        pht_fwant = str(PHT_FWANT)
                        pht_swrant = str(PHT_SWRANT)

                        kpo_fwant = str(KPO_FWANT)
                        kpo_swrant = str(KPO_SWRANT)

                        kai_fwant = str(KAI_FWANT)
                        kai_swrant = str(KAI_SWRANT)

                        #SFN and Emmision to string
                        sfn_lhs = str(SFN_MARGIN_LHS)
                        sfn_tas = str(SFN_MARGIN_TAS)
                        sfn_kpo = str(SFN_MARGIN_KPO)
                        sfn_kai = str(SFN_MARGIN_KAI)
                        sfn_pnp = str(SFN_MARGIN_PNP)

                        em_lhs = str(EM_LHS)
                        em_tas = str(EM_TAS)
                        em_kpo = str(EM_KPO)
                        em_kai = str(EM_KAI)
                        em_pnp = str(EM_PNP)

                        

                        # filename = "report.xlsx"

                        #workbook = Workbook()
                        # sheet = workbook.active

                        # #CHP ird 
                        # sheet["AA33"] = ird_chp_lm_a
                        # sheet["AB33"] = ird_chp_cn_a
                        # sheet["AC33"] = ird_chp_lm_b
                        # sheet["AD33"] = ird_chp_cn_b

                        # workbook.save(filename=filename)

                        wb = openpyxl.load_workbook(fileContent)

                        #sheet = xfile.wb['Daily report']
                        sheet = wb["Daily report"]
                        sheet["I6"] = ird_chp_lm_a
                        sheet["J6"] = ird_chp_cn_a
                        sheet["K6"] = ird_chp_lm_b
                        sheet["L6"] = ird_chp_cn_b
                        sheet["T6"] = ups_runtime_chp
                        sheet["Q6"] = chp_upper
                        sheet["O6"] = chp_lower
                        sheet["P6"] = chp_fwupper
                        sheet["N6"] = chp_fwlower
                        sheet["R6"] = chp_load_ups

                        

                        sheet["I7"] = ird_rng_lm_a
                        sheet["J7"] = ird_rng_cn_a
                        sheet["K7"] = ird_rng_lm_b
                        sheet["L7"] = ird_rng_cn_b
                        sheet["S7"] = ups_runtime_rng_h
                        sheet["T7"] = ups_runtime_rng_min
                        sheet["Q7"] = rng_upper
                        sheet["O7"] = rng_lower
                        sheet["P7"] = rng_fwupper
                        sheet["N7"] = rng_fwlower
                        sheet["R7"] = rng_load_ups


                        sheet["I12"] = ird_pkk_lm_a
                        sheet["J12"] = ird_pkk_cn_a
                        sheet["K12"] = ird_pkk_lm_b
                        sheet["L12"] = ird_pkk_cn_b
                        sheet["T12"] = ups_runtime_pkk
                        sheet["R12"] = pkk_load_ups

                        sheet["I13"] = ird_tas_lm_a
                        sheet["J13"] = ird_tas_cn_a
                        sheet["K13"] = ird_tas_lm_b
                        sheet["L13"] = ird_tas_cn_b
                        sheet["E13"] = sfn_tas
                        sheet["F13"] = em_tas

                        sheet["I14"] = ird_lhs_lm_a
                        sheet["J14"] = ird_lhs_cn_a
                        sheet["K14"] = ird_lhs_lm_b
                        sheet["L14"] = ird_lhs_cn_b
                        sheet["E14"] = sfn_lhs
                        sheet["F14"] = em_lhs

                        sheet["I8"] = ird_tsk_lm_a
                        sheet["J8"] = ird_tsk_cn_a
                        sheet["K8"] = ird_tsk_lm_b
                        sheet["L8"] = ird_tsk_cn_b
                        sheet["T8"] = ups_runtime_tsk
                        sheet["Q8"] = tsk_upper
                        sheet["O8"] = tsk_lower
                        sheet["P8"] = tsk_fwupper
                        sheet["N8"] = tsk_fwlower
                        sheet["R8"] = tsk_load_ups

                        sheet["I15"] = ird_pbi_lm_a
                        sheet["J15"] = ird_pbi_cn_a
                        sheet["K15"] = ird_pbi_lm_b
                        sheet["L15"] = ird_pbi_cn_b

                        sheet["I16"] = ird_hua_lm_a
                        sheet["J16"] = ird_hua_cn_a
                        sheet["K16"] = ird_hua_lm_b
                        sheet["L16"] = ird_hua_cn_b

                        sheet["I9"] = ird_pht_lm
                        sheet["J9"] = ird_pht_cn
                        sheet["T9"] = ups_runtime_pht
                        sheet["O9"] = pht_swrant
                        sheet["N9"] = pht_fwant

                        sheet["I10"] = ird_kpo_lm
                        sheet["J10"] = ird_kpo_cn
                        sheet["T10"] = ups_runtime_kpo
                        sheet["O10"] = kpo_swrant
                        sheet["N10"] = kpo_fwant
                        sheet["E10"] = sfn_kpo
                        sheet["F10"] = em_kpo

                        sheet["I11"] = ird_kura_lm
                        sheet["J11"] = ird_kura_cn
                        sheet["T11"] = ups_runtime_kura
                        sheet["O11"] = kai_swrant
                        sheet["N11"] = kai_fwant
                        sheet["E11"] = sfn_kai
                        sheet["F11"] = em_kai

                        sheet["I17"] = ird_swi_lm
                        sheet["J17"] = ird_swi_cn

                        sheet["I18"] = ird_pnp_lm
                        sheet["J18"] = ird_pnp_cn
                        sheet["E18"] = sfn_pnp
                        sheet["F18"] = em_pnp

                        sheet["I19"] = ird_kbr_lm
                        sheet["J19"] = ird_kbr_cn
        
                        wb.save('รายงานประจำวัน.xlsx')
		
        except:
	         pass

def run():
     if os.path.isfile('./รายงานประจำวัน.xlsx'):   
        GB = 100
        download = 0
        speed = 1
        while(download<GB):
            time.sleep(0.05)
            bar['value']+=(speed/GB)*100
            download+=speed
            percent.set(str(int((download/GB)*100))+"%")
            text.set(str(download)+"/"+" completed")
            window.update_idletasks()


def save_file():
        user_profile = os.environ['USERPROFILE']
        user_desktop = user_profile + "/Desktop"        
        if os.path.isfile('./รายงานประจำวัน.xlsx'):
                shutil.move('./รายงานประจำวัน.xlsx', user_desktop)


def popupmsg():
 if os.path.isfile('./รายงานประจำวัน.xlsx'):
    popup = tk.Tk()
    popup.wm_title("การแจ้งเตือน")
    label = ttk.Label(popup, text="ไฟล์รายงานของคุณได้ถูกส่งไปยังโฟลเดอร์ Desktop เรียบร้อยแล้ว", font=NORM_FONT)
    label.pack(side="top", fill="x", pady=30)
    B1 = ttk.Button(popup, text="Ok", command = popup.destroy)
    B1.pack()
    popup.mainloop()

LARGE_FONT= ("Verdana", 12)
NORM_FONT = ("Helvetica", 10)
SMALL_FONT = ("Helvetica", 8)

window = Tk()
window.title('โปรแกรมทำรายงานประจำวันอัตโนมัติ V2.0')
window.geometry("1000x500")

# Add image file
bg = PhotoImage(file = "Auto report.png")
  
# Create Canvas
canvas1 = Canvas( window, width = 1000,
                 height = 500)
  
canvas1.pack(fill = "both", expand = True)
  
# Display image
canvas1.create_image( 0, 0, image = bg, 
                     anchor = "nw")

# Create Buttons
button1 = Button( window, text = "เลือกไฟล์",command = lambda:[play(),selectFile(),play(),run()])
button2 = Button( window, text = "ดาวน์โหลดไฟล์ผลลัพธ์",command=lambda:[play(),popupmsg(),save_file()])

#create progressbar
percent = StringVar()
text = StringVar()

bar = Progressbar(window,orient=HORIZONTAL,length=300)
bar.pack(pady=10)

percentLabel = Label(window,textvariable=percent)
taskLabel = Label(window,textvariable=text)

 

# Display Buttons
button1_canvas = canvas1.create_window( 435, 220, 
                                       anchor= "nw",
                                       window = button1)
  
button2_canvas = canvas1.create_window( 410, 280, 
                                       anchor = "nw",
                                       window = button2)

      

bar_canvas  = canvas1.create_window(320, 50  ,anchor = "nw",window = bar) 

percentLabel_canvas = canvas1.create_window(450,110  ,anchor = "nw",window = percentLabel) 

taskLabel_canvas = canvas1.create_window(430, 160  ,anchor = "nw",window = taskLabel)


window.mainloop()