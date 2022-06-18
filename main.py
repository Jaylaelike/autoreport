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
                        chp = requests.get("https://script.google.com/macros/s/AKfycbwH4k8SD5F6duZLthsa2q3D2wtJvpcggY5HtMC479aB1yg91pchlGk65WVQaDVYfLhd_g/exec")
                        print(chp.text)

                        rng = requests.get("https://script.google.com/macros/s/AKfycbzx8ruu5KJiNctKp12OgFI3XP4dfn_ffDdRZMCD8mZmHbzKuYmfKzb7T8-cAz69mYdt-A/exec")
                        print(rng.text)

                        pkk = requests.get("https://script.google.com/macros/s/AKfycbyoX7cM1Ejkeej7N-TKNAR--l6ZlGSBa0FPiy9jFHq0BU3xyMUbRyHoY5s5Wp9nM41ANg/exec")
                        print(pkk.text)

                        hua = requests.get("https://script.google.com/macros/s/AKfycby2cC69FuKvS3-n1pG3WyE95Lj7CL9ut5k17BgYmbCCvjLB0GA8M5E5FU3K8ya8F_Xk/exec")
                        print(hua.text)

                        pnp = requests.get("https://script.google.com/macros/s/AKfycbzsnI5ILJLerqWqLhKAPpgESAVDA8SNAnP3p2vMRT06SI-jD6GsP8bO4IB4_oeDRTs8Gw/exec")
                        print(pnp.text)

                        pbi = requests.get("https://script.google.com/macros/s/AKfycbzv9izLctRhphOIn4OD2HUgQQuttu8lOIfrh6Lod9W5mbtVVlyFQC1cwdP81SEtyH6oXQ/exec")
                        print(pbi.text)

                        tsk = requests.get("https://script.google.com/macros/s/AKfycbyWPDK5oe2RhlSxzDsDhVvooI_N90BmhIFdBengbgO7Aq6ZwAgvqhWdI1EyVfkGl9kW/exec")
                        print(tsk.text)


                        lhs = requests.get("https://script.google.com/macros/s/AKfycbyazX9wgJTKbF34n4z3QDM1fPLoRwtLho5RVWaHfg52gx-ewz4HgXzHzV9Bc-h9Fhcwew/exec")
                        print(lhs.text)

                        swi = requests.get("https://script.google.com/macros/s/AKfycbywAXbCkS_7zTLEv2hUoPHs6aI9RCohwnX94kDb4j2fpomHbuqIUghsd6YImESsn7xdWA/exec")
                        print(swi.text)

                        kbr = requests.get("https://script.google.com/macros/s/AKfycbwCn90h7bXc7L_L570Fx8k_GyEhSUrzo6eKBSrMn8lwYa42EcF1okSGBYiJrkwcZCT5QQ/exec")
                        print(kbr.text)

                        kpo = requests.get("https://script.google.com/macros/s/AKfycbyJi6htdI-uKUd5dsbMQrDO9VhDTuGEFWdLurhZONePg9_TcTHzSX9rCFKn65REIKRl-w/exec")
                        print(kpo.text)

                        pht = requests.get("https://script.google.com/macros/s/AKfycbwJBTdWmfkZs8w3vstxTlvKgVvbfeQ-3PEEgFIrPvlIg9K12E_31BHi3wrVtHZfDbKDUg/exec")
                        print(pht.text)

                        kura = requests.get("https://script.google.com/macros/s/AKfycbzryrF3abgTU_9HllNohAJ9_4bSgkT6QRWasBrqQt0puNbJUGOhAI8xiOMzrUD1w2eN/exec")
                        print(kura.text)

                        tas = requests.get("https://script.google.com/macros/s/AKfycbxRFNFuoNBZj5M5xFX7o0mcraXPKSnGNI5DGHToNy9fAJPSU9_C-KqmdFPwtw1dJaQn/exec")
                        print(tas.text)
                        
                        chp_swr = requests.get("https://script.google.com/macros/s/AKfycbz2uX3RfbkDNguEoQN3K3aen1aM2t-GrUMEoJT-GvYqffSdVVJ0bwYm5uW5z1SHHFHaeQ/exec")
                        print(chp_swr.text)

                        tsk_swr = requests.get("https://script.google.com/macros/s/AKfycbz3iuBRqvBcBnh3_imzbW0uefRUYwZM_TBPWYF_1zXo2FMAPYLQfm9mu6XcZbh6E7wp/exec")
                        print(tsk_swr.text)

                        rng_swr = requests.get("https://script.google.com/macros/s/AKfycby-tNv2gLU-siv93gCtwU-Oc_HgQQ2UNXA3rhkH8zh-LI9BeyumnHYVp1-HR1VdjPgJ/exec")
                        print(rng_swr.text)

                        pht_swr = requests.get("https://script.google.com/macros/s/AKfycby3kmOHaBPIBt_t5XwN_SEtKjssXMvlu1UC89bOSGvt1Wl89ORInBgUeAEC1ohAQbMI/exec")
                        print(pht_swr.text)

                        kpo_swr = requests.get("https://script.google.com/macros/s/AKfycbwTY9-Zj4_2JsL7dW2xlwD9Rne2eyfQmcB4ZFZ-Hn1NN4Z-Y9FYtaTGVsxgJfrvTaSq/exec")
                        print(kpo_swr.text)

                        kai_swr = requests.get("https://script.google.com/macros/s/AKfycbyLksPAsU9lmw8XLASwFZWzLCAwk-g_N3z83Ixg1C-fqDUuSDmwCUpVdA2QcCsf4tQH/exec")
                        print(kai_swr.text)

                        sfn = requests.get("https://script.google.com/macros/s/AKfycbxqCIeBX7-1_tIl1EiIxDN0b_BW_nwi7jJeIlQybyWVi4tE8L1YVyLGqe258ff3gXjsOQ/exec")
                        print(sfn.text)
                        
                        loadUps = requests.get("https://script.google.com/macros/s/AKfycbwviJ3VJCyy0d1sRmOI0AfGasqkzVlEsRmoInCBwds5pvBStrJY3sE-K0JlN4TdQeg21g/exec")
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


                        upDateTime = json_data_chp['DateTime']

                        # #POWER Mux OR A Watt
                        # CHP_PowerW = json_data_chp[0]['PowerMux(W)']  
                        # RNG_PowerW = json_data_rng[0]['PowerMux(W)'] 
                        # PKK_PowerW = json_data_pkk[0]['PowerMux(W)'] 
                        # HUA_PowerW = json_data_hua[0]['PowerMux(W)'] 
                        # PNP_PowerW = json_data_pnp[0]['PowerMux(W)']
                        # PBI_PowerW = json_data_pbi[0]['PowerA(W)']
                        # LHS_PowerW = json_data_lhs[0]['PowerA(W)']
                        # SWI_PowerW = json_data_swi[0]['PowerA(W)']    
                        # KBR_PowerW = json_data_kbr[0]['Power(W)']
                        # KPO_PowerW = json_data_kpo[0]['Power(W)']
                        # PHT_PowerW = json_data_pht[0]['Power(W)']
                        # KURA_PowerW = json_data_kura[0]['Power(W)']
                        # TAS_PowerW = json_data_tas[0]['PowerA(W)']

                        # #POWER B Watt
                        # PBI_PowerBW = json_data_pbi[0]['PowerB(W)']
                        # LHS_PowerBW = json_data_lhs[0]['PowerB(W)']
                        # TAS_PowerBW = json_data_tas[0]['PowerB(W)']

                        #SFN
                        SFN_MARGIN_LHS = json_data_sfn['SFN_LHS']
                        SFN_MARGIN_TAS = json_data_sfn['SFN_TAS']
                        SFN_MARGIN_KPO = json_data_sfn['SFN_KPO']
                        SFN_MARGIN_KAI = json_data_sfn['SFN_KAI']
                        SFN_MARGIN_PNP = json_data_sfn['SFN_PNP']
                        SFN_MARGIN_KBR = json_data_sfn['SFN_KBR']
                        SFN_MARGIN_PHT = json_data_sfn['SFN_PHT']
                        SFN_MARGIN_SWI = json_data_sfn['SFN_SWI']

                        #SFN_NEC is string from API
                        sfn_chp_mux = str(json_data_sfn['SFN_CHP_MUX'])
                        ems_chp_mux = str(json_data_sfn['EMS_CHP_MUX'])
                        sfn_chp_res = str(json_data_sfn['SFN_CHP_RES'])
                        ems_chp_res = str(json_data_sfn['EMS_CHP_RES'])
                        sfn_rng_mux = str(json_data_sfn['SFN_RNG_MUX'])
                        ems_rng_mux = str(json_data_sfn['EMS_RNG_MUX'])
                        sfn_rng_res = str(json_data_sfn['SFN_RNG_RES']) 
                        ems_rng_res = str(json_data_sfn['EMS_RNG_RES'])
                        sfn_pkk_mux = str(json_data_sfn['SFN_PKK_MUX'])  
                        ems_pkk_mux = str(json_data_sfn['EMS_PKK_MUX'])
                        sfn_pkk_res = str(json_data_sfn['SFN_PKK_RES'])
                        ems_pkk_res = str(json_data_sfn['EMS_PKK_RES'])
                        sfn_hua_mux = str(json_data_sfn['SFN_HUA_MUX']) 
                        ems_hua_mux = str(json_data_sfn['EMS_HUA_MUX'])
                        sfn_hua_res = str(json_data_sfn['SFN_HUA_RES']) 
                        ems_hua_res = str(json_data_sfn['EMS_HUA_RES'])



                        #Emmission time
                        EM_LHS = json_data_sfn['EMMIS_LHS']
                        EM_TAS = json_data_sfn['EMMIS_TAS']
                        EM_KPO = json_data_sfn['EMMIS_KPO']
                        EM_KAI = json_data_sfn['EMMIS_KAI']
                        EM_PNP = json_data_sfn['EMMIS_PNP']
                        EM_KBR = json_data_sfn['EMMIS_KBR']
                        EM_PHT = json_data_sfn['EMMIS_PHT']
                        EM_SWI = json_data_sfn['EMMIS_SWI']



                        #SWR
                        CHP_UPPER   = json_data_chp_swr['upper']
                        CHP_LOWER   = json_data_chp_swr['lower']
                        CHP_FWLOWER = json_data_chp_swr['FWD_lower']
                        CHP_FWUPPER = json_data_chp_swr['FWD_upper']

                        TSK_UPPER = json_data_tsk_swr['upper']
                        TSK_LOWER   = json_data_tsk_swr['lower']
                        TSK_FWLOWER = json_data_tsk_swr['FWD_lower']
                        TSK_FWUPPER = json_data_tsk_swr['FWD_upper']

                        RNG_UPPER = json_data_rng_swr['upper']
                        RNG_LOWER   = json_data_rng_swr['lower']
                        RNG_FWLOWER = json_data_rng_swr['FWD_lower']
                        RNG_FWUPPER = json_data_rng_swr['FWD_upper']

                        PHT_FWANT = json_data_pht_swr['FWD_ant']
                        PHT_SWRANT = json_data_pht_swr['SWR_ant']

                        KPO_FWANT = json_data_kpo_swr['FWD_ant']
                        KPO_SWRANT = json_data_kpo_swr['SWR_ant']

                        KAI_FWANT = json_data_kai_swr['FWD_ant']
                        KAI_SWRANT = json_data_kai_swr['SWR_ant']


                        #IRD
                        CHP_IRD_CN_A = json_data_chp['IRD_A_C/N'] 
                        CHP_IRD_LM_A = json_data_chp['IRD_A_LM']
                        CHP_IRD_CN_B = json_data_chp['IRD_B_C/N'] 
                        CHP_IRD_LM_B = json_data_chp['IRD_B_LM']

                        RNG_IRD_CN_A = json_data_rng['IRD_A_C/N'] 
                        RNG_IRD_LM_A = json_data_rng['IRD_A_LM']
                        RNG_IRD_CN_B = json_data_rng['IRD_B_C/N'] 
                        RNG_IRD_LM_B = json_data_rng['IRD_B_LM']

                        PKK_IRD_CN_A = json_data_pkk['IRD_A_C/N'] 
                        PKK_IRD_LM_A = json_data_pkk['IRD_A_LM'] 
                        PKK_IRD_CN_B = json_data_pkk['IRD_B_C/N'] 
                        PKK_IRD_LM_B = json_data_pkk['IRD_B_LM']

                        HUA_IRD_CN_A = json_data_hua['IRD_A_C/N'] 
                        HUA_IRD_LM_A = json_data_hua['IRD_A_LM']
                        HUA_IRD_CN_B = json_data_hua['IRD_B_C/N'] 
                        HUA_IRD_LM_B = json_data_hua['IRD_B_LM']

                        TSK_IRD_CN_A = json_data_tsk['IRD_A_CN'] 
                        TSK_IRD_LM_A = json_data_tsk['IRD_A_LM']
                        TSK_IRD_CN_B = json_data_tsk['IRD_B_CN'] 
                        TSK_IRD_LM_B = json_data_tsk['IRD_B_LM']

                        PNP_IRD_CN = json_data_pnp['IRD_A_CN'] 
                        PNP_IRD_LM = json_data_pnp['IRD_A_LM']

                        PBI_IRD_CN_A = json_data_pbi['IRD_A_CN'] 
                        PBI_IRD_LM_A = json_data_pbi['IRD_A_LM'] 
                        PBI_IRD_CN_B = json_data_pbi['IRD_B_CN'] 
                        PBI_IRD_LM_B = json_data_pbi['IRD_B_LM'] 

                        LHS_IRD_CN_A = json_data_lhs['IRD_A_CN'] 
                        LHS_IRD_LM_A = json_data_lhs['IRD_A_LM']
                        LHS_IRD_CN_B = json_data_lhs['IRD_B_CN'] 
                        LHS_IRD_LM_B = json_data_lhs['IRD_B_LM'] 

                        SWI_IRD_CN = json_data_swi['IRD_A_CN'] 
                        SWI_IRD_LM = json_data_swi['IRD_A_LM']  

                        KBR_IRD_CN = json_data_kbr['IRD_A_CN'] 
                        KBR_IRD_LM = json_data_kbr['IRD_A_LM']

                        KPO_IRD_CN = json_data_kpo['IRD_A_CN'] 
                        KPO_IRD_LM = json_data_kpo['IRD_A_LM']

                        PHT_IRD_CN = json_data_pht['IRD_A_CN'] 
                        PHT_IRD_LM = json_data_pht['IRD_A_LM']  

                        KURA_IRD_CN = json_data_kura['IRD_A_CN'] 
                        KURA_IRD_LM = json_data_kura['IRD_A_LM']  

                        TAS_IRD_CN_A = json_data_tas['IRD_A_CN'] 
                        TAS_IRD_LM_A = json_data_tas['IRD_A_LM'] 
                        TAS_IRD_CN_B = json_data_tas['IRD_B_CN'] 
                        TAS_IRD_LM_B = json_data_tas['IRD_B_LM']  
                        

                        # #UPS
                        CHP_UpsRuntime = json_data_chp['UPSRuntime']
                        RNG_UpsRuntime = json_data_rng['UPSRuntime']
                        PKK_UpsRuntime = json_data_pkk['UPSRuntime']
                        TSK_UpsRuntime = json_data_tsk['UPSRuntime']
                        PHT_UpsRuntime = json_data_pht['UPS_Runtime']
                        KPO_UpsRuntime = json_data_kpo['UPS_Runtime']
                        KURA_UpsRuntime = json_data_kura['UPS_Runtime']

                        # load ups %
                        CHP_load_ups = json_data_loadUps['loadCHP']
                        RNG_load_ups = json_data_loadUps['loadRNG']
                        PKK_load_ups = json_data_loadUps['loadPKK']
                        TSK_load_ups = json_data_loadUps['loadTSK']
                        PHT_load_ups = json_data_loadUps['loadPHT']
                        KPO_load_ups = json_data_loadUps['loadKPO']
                        KAI_load_ups = json_data_loadUps['loadKAI']


                



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
                        ups_runtime_chp = str(int(CHP_UpsRuntime)) 
                        ups_runtime_rng_h = str(rngTimehours)
                        ups_runtime_rng_min = str(rngTimeMinutes)
                        ups_runtime_pkk = str(int(PKK_UpsRuntime))
                        ups_runtime_pht = str(int(PHT_UpsRuntime))
                        ups_runtime_kpo = str(int(KPO_UpsRuntime))
                        ups_runtime_kura = str(int(KURA_UpsRuntime))
                        ups_runtime_tsk = str(int(TSK_UpsRuntime))

                        #load ups % to string
                        chp_load_ups = str(CHP_load_ups)
                        rng_load_ups = str(RNG_load_ups)
                        pkk_load_ups = str(PKK_load_ups)
                        tsk_load_ups = str(TSK_load_ups)
                        pht_load_ups = str(PHT_load_ups)
                        kpo_load_ups = str(KPO_load_ups)
                        kai_load_ups = str(KAI_load_ups)



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
                        sfn_kbr = str(SFN_MARGIN_KBR)
                        sfn_pht = str(SFN_MARGIN_PHT)
                        sfn_swi = str(SFN_MARGIN_SWI)


                        em_lhs = str(EM_LHS)
                        em_tas = str(EM_TAS)
                        em_kpo = str(EM_KPO)
                        em_kai = str(EM_KAI)
                        em_pnp = str(EM_PNP)
                        em_kbr = str(EM_KBR)
                        em_pht = str(EM_PHT)
                        em_swi = str(EM_SWI)


                        

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
                        sheet = wb["CHP"]
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
                        sheet["G14"] = sfn_lhs   #B
                        sheet["H14"] = em_lhs    #B 

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
                        sheet["E9"] = sfn_pht
                        sheet["F9"] = em_pht
                        sheet["R9"] = pht_load_ups


                        sheet["I10"] = ird_kpo_lm
                        sheet["J10"] = ird_kpo_cn
                        sheet["T10"] = ups_runtime_kpo
                        sheet["O10"] = kpo_swrant
                        sheet["N10"] = kpo_fwant
                        sheet["E10"] = sfn_kpo
                        sheet["F10"] = em_kpo
                        sheet["R10"] = kpo_load_ups
                        

                        sheet["I11"] = ird_kura_lm
                        sheet["J11"] = ird_kura_cn
                        sheet["T11"] = ups_runtime_kura
                        sheet["O11"] = kai_swrant
                        sheet["N11"] = kai_fwant
                        sheet["E11"] = sfn_kai
                        sheet["F11"] = em_kai
                        sheet["R11"] = kai_load_ups

                        sheet["I17"] = ird_swi_lm
                        sheet["J17"] = ird_swi_cn
                        sheet["E17"] = sfn_swi
                        sheet["F17"] = em_swi

                        sheet["I18"] = ird_pnp_lm
                        sheet["J18"] = ird_pnp_cn
                        sheet["E18"] = sfn_pnp
                        sheet["F18"] = em_pnp

                        sheet["I19"] = ird_kbr_lm
                        sheet["J19"] = ird_kbr_cn
                        sheet["E19"] = sfn_kbr
                        sheet["F19"] = em_kbr

                        #NEC_SFN_EMSIION TIME
                        sheet["E6"] = sfn_chp_mux
                        sheet["F6"] = ems_chp_mux
                        sheet["G6"] = sfn_chp_res
                        sheet["H6"] = ems_chp_res
                        sheet["E7"] = sfn_rng_mux
                        sheet["F7"] = ems_rng_mux
                        sheet["G7"] = sfn_rng_res
                        sheet["H7"] = ems_rng_res
                        sheet["E12"] = sfn_pkk_mux
                        sheet["F12"] = ems_pkk_mux
                        sheet["G12"] = sfn_pkk_res
                        sheet["H12"] = ems_pkk_res
                        sheet["E16"] = sfn_hua_mux
                        sheet["F16"] = ems_hua_mux
                        sheet["G16"] = sfn_hua_res
                        sheet["H16"] = ems_hua_res


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
window.title('โปรแกรมทำรายงานประจำวันอัตโนมัติ V5.0')
window.geometry("1000x500")

# Add image file
bg = PhotoImage(file = "Auto report2.png")
  
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