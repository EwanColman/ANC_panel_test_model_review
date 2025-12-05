# -*- coding: utf-8 -*-
"""
Created on Mon Nov 10 16:44:47 2025

@author: iv23655
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

average_over_time=False

# dictionary of all variables
v={}

# Country data
df=pd.read_excel('Abbott YHEC ANC panel CEA and BIA FINAL_V6.0_25.11.2025_Nigeria.xlsm',
                 sheet_name='Country data',
                 skiprows=16,
                 usecols='D:E',
                 index_col=0,
                 nrows=322-16)

for i,row in df.iterrows():
    v[i]=row['Nigeria']

# live data 
df=pd.read_excel('Abbott YHEC ANC panel CEA and BIA FINAL_V6.0_25.11.2025_Nigeria.xlsm',
                 sheet_name='Live sheet',
                 skiprows=38,
                 usecols='D:F',
                 index_col=0,
                 nrows=314-16)

for i,row in df.iterrows():
    v[i]=row['Live values']



# some random extra variables
df=pd.read_excel('Abbott YHEC ANC panel CEA and BIA FINAL_V6.0_25.11.2025_Nigeria.xlsm',
                 sheet_name='Live sheet',
                 skiprows=23,
                 usecols='G:I',
                 index_col=0,
                 nrows=3,
                 names=['Variable','blank','Value']
                 )

# "minutes in an hour"? why is this variable? 
for i,row in df.iterrows():
    v[i]=row['Value']

# Gestational parent general population HRQoL and mortality probability
gest_parent_utility=pd.read_excel('Abbott YHEC ANC panel CEA and BIA FINAL_V6.0_25.11.2025_Nigeria.xlsm',
                 sheet_name='Live sheet',
                 skiprows=322,
                 usecols='D:O',
                 index_col=0,
                 nrows=4
                 )


##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~##

N={}

## Sheet: Set up
# E42
cohort=v['Cohort size']#v['Percentage of pregnant people seeking ANC ']
# different in "Live values" compared to "Country data"

## Sheet: ANC Acute HIV
#  G76
N['Acute_HIV']=cohort*v['Prevalence of Acute HIV in pregnant people']

# G115
N['No_acute_HIV']=cohort*(1-v['Prevalence of Acute HIV in pregnant people'])

#~~~~#

# J65
N[('Acute_HIV','HIV_PoC_Test')]=N['Acute_HIV']

# J88
N[('Acute_HIV','No_HIV_PoC_test')]=0 # why have other descendant states if always 0?
# assumption seems to be that 100% receive ANC test, but its less than that for SoC
# Q: does getting a SoC PoC test depend on disease status
    
# J108
N[('No_acute_HIV','HIV_PoC_Test')]=N['No_acute_HIV']

# J88
N[('No_acute_HIV','No_HIV_PoC_test')]=0 # why have other descendant states if always 0?

#~~~~#

# M55
N[('Acute_HIV','HIV_PoC_Test','True positive')]=N[('Acute_HIV','HIV_PoC_Test')]*v['Acute HIV sensitivity (%) of Determine Antenatal Care Panel']

# M76
N[('Acute_HIV','HIV_PoC_Test','False negative')]=N[('Acute_HIV','HIV_PoC_Test')]*(1-v['Acute HIV sensitivity (%) of Determine Antenatal Care Panel'])

# M102
N[('No_acute_HIV','HIV_PoC_Test','False positive')]=N[('No_acute_HIV','HIV_PoC_Test')]*(1-v['Acute HIV specificity (%) of Determine Antenatal Care Panel'])

# M114
N[('No_acute_HIV','HIV_PoC_Test','True negative')]=N[('No_acute_HIV','HIV_PoC_Test')]*v['Acute HIV specificity (%) of Determine Antenatal Care Panel']

#~~~~#

# P46
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test')]=N[('Acute_HIV','HIV_PoC_Test','True positive')]*(1-v['Acute HIV: percent lost to follow-up (%) after Determine Antenatal Care Panel'])
# param set to 0, 100% getting conformatory test

# P64
N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up')]=N[('Acute_HIV','HIV_PoC_Test','True positive')]*v['Acute HIV: percent lost to follow-up (%) after Determine Antenatal Care Panel']
 
# P98
N[('No_acute_HIV','HIV_PoC_Test','False positive','Confirmatory test')]=N[('No_acute_HIV','HIV_PoC_Test','False positive')]*(1-v['Acute HIV: percent lost to follow-up (%) after Determine Antenatal Care Panel'])

# P106
N[('No_acute_HIV','HIV_PoC_Test','False positive','Lost to follow-up')]=N[('No_acute_HIV','HIV_PoC_Test','False positive')]*v['Acute HIV: percent lost to follow-up (%) after Determine Antenatal Care Panel']

#~~~~#

# S40
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test')]*v['Percentage (%) of pregnant people receiving treatment for Acute HIV']
# Would you have false negatives for those already receiving treatment? What are the ART assumptions here?
# General question where do pregnant woment with known HIV+ status fit in to this model?

# S52
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test')]*(1-v['Percentage (%) of pregnant people receiving treatment for Acute HIV'])

#~~~~#

# V36
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART','Perinatal mortality')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART')]*v['Probability (%) of Acute HIV-related perinatal fetal mortality (receiving prenatal treatment)']

# V40
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART','HIV positive infant')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (receiving prenatal treatment)'])*v['Probability of vertical Acute HIV transmission (%) with treatment']

# V44
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART','Normal full-term delivery')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (receiving prenatal treatment)'])*(1-v['Probability of vertical Acute HIV transmission (%) with treatment'])

# V48
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access','Perinatal mortality')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access')]*v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)']

# V52 
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access','HIV positive infant')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*v['Probability of vertical Acute HIV transmission (%) without treatment']

# V56
N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access','Normal full-term delivery')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','No treatment access')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*(1-v['Probability of vertical Acute HIV transmission (%) without treatment'])

# V60
N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up','Perinatal mortality')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up')]*v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)']
# Looks like true positives lost to follow up can't be put on treatment - cohort assumed fully HIV-?

# V64 
N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up','HIV positive infant')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*v['Probability of vertical Acute HIV transmission (%) without treatment']

# V68 
N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up','Normal full-term delivery')]=N[('Acute_HIV','HIV_PoC_Test','True positive','Lost to follow-up')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*(1-v['Probability of vertical Acute HIV transmission (%) without treatment'])

# V72
N[('Acute_HIV','HIV_PoC_Test','False negative','Perinatal mortality')]=N[('Acute_HIV','HIV_PoC_Test','False negative')]*v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)']

# V76
N[('Acute_HIV','HIV_PoC_Test','False negative','HIV positive infant')]=N[('Acute_HIV','HIV_PoC_Test','False negative')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*v['Probability of vertical Acute HIV transmission (%) without treatment']

# V80
N[('Acute_HIV','HIV_PoC_Test','False negative','Normal full-term delivery')]=N[('Acute_HIV','HIV_PoC_Test','False negative')]*(1-v['Probability (%) of Acute HIV-related perinatal fetal mortality (not receiving prenatal treatment)'])*(1-v['Probability of vertical Acute HIV transmission (%) without treatment'])

# V96
N[('No_acute_HIV','HIV_PoC_Test','False positive','Confirmatory test','Perinatal mortality')]=N[('No_acute_HIV','HIV_PoC_Test','False positive','Confirmatory test')]*v['Probability (%) of perinatal fetal mortality in the general population']

# V100
N[('No_acute_HIV','HIV_PoC_Test','False positive','Confirmatory test','Normal full-term delivery')]=N[('No_acute_HIV','HIV_PoC_Test','False positive','Confirmatory test')]*(1-v['Probability (%) of perinatal fetal mortality in the general population'])

# V104
N[('No_acute_HIV','HIV_PoC_Test','False positive','Lost to follow-up','Perinatal mortality')]=N[('No_acute_HIV','HIV_PoC_Test','False positive','Lost to follow-up')]*v['Probability (%) of perinatal fetal mortality in the general population']

# V108
N[('No_acute_HIV','HIV_PoC_Test','False positive','Lost to follow-up','Normal full-term delivery')]=N[('No_acute_HIV','HIV_PoC_Test','False positive','Lost to follow-up')]*(1-v['Probability (%) of perinatal fetal mortality in the general population'])

# V112
N[('No_acute_HIV','HIV_PoC_Test','True negative','Perinatal mortality')]=N[('No_acute_HIV','HIV_PoC_Test','True negative')]*v['Probability (%) of perinatal fetal mortality in the general population']

# V116
N[('No_acute_HIV','HIV_PoC_Test','True negative','Normal full-term delivery')]=N[('No_acute_HIV','HIV_PoC_Test','True negative')]*(1-v['Probability (%) of perinatal fetal mortality in the general population'])

for key,item in N.items():
    print(key,'\t',item)
print()
####### COSTS ##############

# dictionaries for time and costs
t,c={},{}

# Z13 and Z19
c['PoC test costs']=v['Cost ($USD) of the Determine Antenatal Care Panel']/v['Number of diseases']
# hcw_time_confirm_test listed as lab_tech_time_confirm_test in Live sheet
    
# AB13 
t['Healthcare worker time']=v['Staff time requirements (minutes) for a healthcare worker to administer the Determine Antenatal Care Panel']/v['Number of diseases']

# AB19
c['Healthcare worker time']=t['Healthcare worker time']/60


# AE18
mean_AHD_QALY_multiplier_with_treatment=v['Chronic HIV-specific utility multipliers for people recieving treatment']*(1-v['Percentage of the HIV pregnant population with AHD at diagnosis'])+v['AHD-specific utility multipliers for people recieving treatment']*v['Percentage of the HIV pregnant population with AHD at diagnosis']
# AE19
mean_AHD_QALY_multiplier_without_treatment=v['Chronic HIV-specific utility multipliers for people not recieving treatment']*(1-v['Percentage of the HIV pregnant population with AHD at diagnosis'])+v['AHD-specific utility multipliers for people not recieving treatment']*v['Percentage of the HIV pregnant population with AHD at diagnosis']
# AG18
mean_AHD_DALY_multiplier_with_treatment=v['Chronic HIV-specific disability weights for people recieving treatment']*(1-v['Percentage of the HIV pregnant population with AHD at diagnosis'])+v['AHD-specific disability weights for people recieving treatment']*v['Percentage of the HIV pregnant population with AHD at diagnosis']
# AG19
mean_AHD_DALY_multiplier_without_treatment=v['Chronic HIV-specific disability weights for people not recieving treatment']*(1-v['Percentage of the HIV pregnant population with AHD at diagnosis'])+v['AHD-specific disability weights for people recieving treatment']*v['Percentage of the HIV pregnant population with AHD at diagnosis']
# these weights/utilities are all for Chronic. No data for Acute?

### AHD weights ###
# AM 13
AHD_weights_with_treatment=[v['Percentage of the HIV pregnant population with AHD at diagnosis']+(1-v['Percentage of the HIV pregnant population with AHD at diagnosis'])*v['Annual probability (%) of Chronic HIV progressing to AHD in gestational parent with treatment']]
for i in range(1,10):
    #AN13:AV13
    AHD_weights_with_treatment.append(AHD_weights_with_treatment[-1]+(1-AHD_weights_with_treatment[-1])*v['Annual probability (%) of Chronic HIV progressing to AHD in gestational parent with treatment'])

# Overall survival
Overall_survival_with_treatment=[]
for i in range(9):
    #AM18:
    Overall_survival_with_treatment.append(gest_parent_utility[i+1].loc['Annual probability of death']*(v['Relative risk of HIV-related gestational parent mortality (not receiving treatment)']*(1-AHD_weights_with_treatment[i])+v['Relative risk of AHD-related gestational parent mortality (not receiving treatment)']*AHD_weights_with_treatment[i]))
    # AM18 is this a mistake? Without treatment used up to year 10. Should "with treatment" variabes be used?
Overall_survival_with_treatment.append(gest_parent_utility[10].loc['Annual probability of death']*(v['Relative risk of HIV-related gestational parent mortality (receiving treatment)']*(1-AHD_weights_with_treatment[9])+v['Relative risk of AHD-related gestational parent mortality (receiving treatment)']*AHD_weights_with_treatment[9]))


# DALYs
DALYs=[]
for i in range(10):
    # AX18:BG18
    
    # fraction who are AHD freezes at year 5 (a mistake I think)
    DALYs.append(v['Chronic HIV-specific disability weights for people recieving treatment']*(1-AHD_weights_with_treatment[min(i,4)])+v['AHD-specific disability weights for people recieving treatment']*AHD_weights_with_treatment[min(i,4)])

# QALYs
QALYs=[]
for i in range(10):
    # BI18:BR18
    # fraction who are AHD freezes at year 5 (a mistake I think)
    QALYs.append(v['Chronic HIV-specific utility multipliers for people recieving treatment']*(1-AHD_weights_with_treatment[min(i,4)])+v['AHD-specific utility multipliers for people recieving treatment']*AHD_weights_with_treatment[min(i,4)])

# QALYs
Costs=[]
for i in range(10):
    # BT18:CC18
    Costs.append(v['Annual cost ($USD) for gestational parent with HIV receiving treatment']*(1-AHD_weights_with_treatment[i])+v['Annual cost ($USD) for  gestational parent with AHD receiving treatment']*AHD_weights_with_treatment[i])


category=('Acute_HIV','HIV_PoC_Test','True positive','Confirmatory test','On-HAART','Perinatal mortality')
c[category]={}

# X36
c[category]['PoC Test cost']=N[category]*c['PoC test costs']
#c[category]['Cost of consumables']=N[category]*0
# Z36
c[category]['Staff time (hours)']=N[category]*c['Healthcare worker time']
# AA36
c[category]['Cost of staff time']=c[category]['Staff time (hours)']*v['Per-hour staff cost ($USD) of a healthcare worker ']
# AB36
c[category]['Confirmatory test cost']=N[category]*(v['Cost ($USD) of the Chronic HIV confirmatory test']+v['Staff time requirements (minutes) for a healthcare worker to administer all confirmatory tests']*v['Per-hour staff cost ($USD) of a healthcare worker ']/60)
# AC36
c[category]['Prenatal treatment cost']=N[category]*v['Cost ($USD) of prenatal treatment for Chronic HIV']
# AD36
c[category]['Intrapartum costs']=N[category]*(v['Cost ($USD) associated with still birth']+v['Cost ($USD) of the intrapartum period'])
# AF36
c[category]['LYs']=N[category]*(1-v['Probability (%) of Acute HIV-related intrapartum gestational parent mortality (receiving prenatal treatment)'])
# AE36
c[category]['QALYs']=c[category]['LYs']*(gest_parent_utility[1].loc['Population norm']*mean_AHD_QALY_multiplier_with_treatment-v['Pregnancy utility decrement']-v['Perinatal death utility decrement (annual)'])
# what is pregnancy utility decrement?
# AG36
c[category]['DALYs']=N[category]-c[category]['LYs']+c[category]['LYs']*mean_AHD_DALY_multiplier_with_treatment


# Societal costs over gestational period
# AI13 and AI19
c['Waiting time']=(v['Time (minutes) taken for pregnant person to travel to an ANC centre']+v['Waiting time (minutes) for the result of the Determine Antenatal Care Panel'])/v['Number of diseases']

# AI36
c[category]['Time for pregnant person']=N[category]*(c['Waiting time']+v['Waiting time (minutes) for the result of the confirmatory Acute HIV test'])/60
# AJ36
c[category]['Income lost']=c[category]['Time for pregnant person']*v['Cost of patient time (per hour)']
# maternity leave mentioned but not factored into calculation??


# AM36
c[category]['Benefits: overall survival TPP']=[c[category]['LYs']*(1-Overall_survival_with_treatment[0])]
for i in range(1,10):
    # AN36:AT36
    c[category]['Benefits: overall survival TPP'].append(c[category]['Benefits: overall survival TPP'][-1]*(1-Overall_survival_with_treatment[i]))

# DALY Benefits
c[category]['Benefits: DALYs TPP']=[]
for i in range(10):
    # AX36:BG36
    c[category]['Benefits: DALYs TPP'].append(c[category]['Benefits: overall survival TPP'][i]*DALYs[i]+(N[category]-c[category]['Benefits: overall survival TPP'][i]))

# QALY benefits
c[category]['Benefits: QALYs TPP']=[]
for i in range(10):
    # BI36:BR36
    c[category]['Benefits: QALYs TPP'].append((gest_parent_utility[i+1].loc['Population norm']*QALYs[i]-v['Perinatal death utility decrement (annual)'])*c[category]['Benefits: overall survival TPP'][i])
    # stillbirth reduction in utility applied every year - is this legit? - check source 

# QALY benefits
c[category]['Costs TPP']=[]
for i in range(10):
    # BT36:CC36
    c[category]['Costs TPP'].append(c[category]['Benefits: overall survival TPP'][i]*Costs[i])


        
for key,item in c[category].items():
    print(key,'\t',item)
    




# dlb_setup_gestation = v['Length of gestational period (years)']    
