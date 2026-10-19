#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
spx_additions_analysis.py: the S&P 500 index-inclusion strategy, from the raw data to the paper's tables
=========================================================================================================

LBS Investment Fundamentals, group paper, appendix code.

The strategy. Buy each stock S&P announces it will add to the S&P 500 at the open of the first trading day after
the announcement (Ann+1). Window A sells at the open of the inclusion day (the first day in the index), Window B at
its close. The benchmark is SPY bought and sold at exactly the same moments. Sample: 278 additions announced from
January 2010 to September 2026, on 206 announcement dates.

The script needs nothing but itself, an internet connection and a WRDS account. It
  1. downloads the data: daily prices of each addition and of SPY from CRSP (through WRDS, which asks for your
     login), prices after CRSP's last day used from Yahoo Finance, and the daily Fama-French factors from
     Kenneth French's website (PART 1);
  2. recomputes every test of the paper from those prices (PART 2). The calculations copy the group's Excel
     workbook (SPX Additions v13.xlsx) formula by formula, and the comments name the workbook sheet and column
     each step reproduces;
  3. prints every table of the paper and the data behind each figure (PART 3).

How to run it in Google Colab
-----------------------------
    1. Open a new notebook at https://colab.research.google.com
    2. Upload this file: click the folder icon on the left, then the upload icon.
    3. In a code cell, run:           !pip install wrds yfinance
       If Colab then asks to restart the session, click Restart session (the uploaded file stays).
    4. In a new code cell, run:       %run spx_additions_analysis.py
    5. Type your WRDS username and password when asked and approve the Duo push on your phone. If WRDS offers to
       create a .pgpass file, answer n.
The run takes a few minutes, most of it downloading.

On your own computer: pip install numpy pandas scipy wrds yfinance, then: python spx_additions_analysis.py
Options (also after %run):
    --details        also print every statistic of every test, laid out like the workbook's result sheets
    --save FOLDER    also save the workings as CSV files: one row per trade, per day a position is held and per
                     trading day of the portfolio

The data
--------
The 278 additions, with their announcement and inclusion dates, were taken by hand from S&P Dow Jones Indices'
press releases ("... Set to Join S&P 500") and are typed into this file (ADDITIONS below), each with its permanent
CRSP number (PERMNO), so a reused ticker cannot pull in another firm's prices. To give exactly the paper's numbers,
the script uses the data as they stood when the paper's data were collected: CRSP prices to 31 December 2025
(CRSP's last day at the time), Yahoo Finance after that, and French's factors to 31 August 2026. The providers
occasionally revise their history, which can move a number in its last decimal; the "Data check" line of the
printout says whether the downloaded data are exactly the paper's. No login details are stored anywhere: the wrds
package asks for them when it connects.

Conventions copied from the spreadsheet (so the numbers agree with the workbook to about 1e-12): STDEV.S divides by
n - 1; RANK.AVG gives tied values their average rank; PERCENTILE.INC is numpy's default (linear) percentile;
T.DIST.2T truncates its degrees of freedom to an integer; the Wilcoxon and Mann-Whitney z use no tie or continuity
correction; "each date once" means the average of the trades announced that day, so G = number of announcement
dates. Printed numbers are rounded half up, as Excel displays them.
"""
import argparse
import datetime as dt
import io
import math
import re
import sys
import time
import urllib.request
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pandas as pd
from scipy import stats

# =====================================================================================================================
# SETTINGS: every assumption of the analysis (the workbook's Settings sheet)
# =====================================================================================================================
ALPHA = 0.05                                  # significance level, two-sided
ZCRIT = stats.norm.ppf(1 - ALPHA / 2)         # 1.96: Figure 1 band and Test 3's "t > 1.96" rule
POWER = 0.80                                  # power for the minimum detectable effect (MDE)
ZPOW = stats.norm.ppf(POWER)                  # 0.84
MDE_MULT = ZCRIT + ZPOW                       # MDE = 2.80 x clustered SE
MIN_EFFECT = 0.01                             # smallest effect worth detecting, per trade (1%)
DAYS_PER_YEAR = 252                           # annualisation
MIN_BETA_DAYS = 100                           # fewer stock returns than this: beta = 1
NW_LAGS = 5                                   # Newey-West lags (Bartlett weights)
SPLIT_YEAR = 2017                             # first half = announced in or before this year
WINSOR_LO, WINSOR_HI = 0.01, 0.99             # V3 winsorising percentiles
TRIM_N = 5                                    # V6: trades dropped at each end
EXCL_YEAR_1, EXCL_YEAR_2 = 2020, 2021         # V4: years left out
#                                               FF_LAST, the last day with Fama-French factors, is read from
#                                               the data (31 Aug 2026): the regressions end there.
START_2020 = dt.date(2020, 1, 1)              # Extra Checks E3: trades announced on or after this day
BREAK_TRIM = 0.15                             # E4: share of dates kept out at each end of the break search
BAI_PERRON_CV_5 = 8.58                        # E4: sup-F critical value at 5% (Bai and Perron 1998, Table I:
BAI_PERRON_CV_1 = 12.29                       #     one break, one regressor, 15% trimming) and at 1%
N_DROP_V1 = 3                                 # E5: largest quarterly trades left out

BETA_FIRST_DAY, BETA_LAST_DAY = 250, 30       # beta estimated over trading days -250 to -30 before the announcement
ACCOUNT_START = dt.date(2010, 1, 1)           # the portfolio starts on the first trading day of 2010 ($1 at the 31 Dec 2009
                                              # close)
HIST_DAYS = 260                               # trading days of prices kept before each announcement
WORKED_EXAMPLE = "Etsy (2020)"                # the trade shown step by step in Tests 1 and 2 (option --details)

# The data (Part 1)
SPY_PERMNO = 84398                            # SPY's permanent number in CRSP
CRSP_END = pd.Timestamp("2025-12-31")         # last CRSP day used; later prices come from Yahoo Finance
FF_END = pd.Timestamp("2026-08-31")           # last Fama-French day used
FIRST_DAY = pd.Timestamp("2008-01-01")        # SPY and the factors are downloaded from this day
FRENCH = "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
# The data behind the paper (SPX Additions v13.xlsx), summed by source. The printout checks a fresh download against
# these sums, so it shows whether a data provider has revised its history since the paper's data were collected.
PAPER_DATA = {"CRSP stock prices": 9315224.738278056, "Yahoo stock prices": 692897.8882026672,
              "CRSP SPY prices": 21423767.85507, "Yahoo SPY prices": 1133227.5786132812, "Fama-French factors": 2.5467}

# =====================================================================================================================
# PART 1. THE DATA: download the prices and factors and lay them out, day by day, for each addition
# =====================================================================================================================
# The 278 additions studied, newest first (the order of the workbook's sheets): every addition to the S&P 500
# announced from 1 January 2010 to 30 September 2026, from S&P Dow Jones Indices' press releases, less 35 spinoffs,
# 22 merger-linked additions and 6 second share classes (63 in all, out of 341).
# announced: the day S&P announced the addition (always after the close); included: the first day in the index;
# permno: the stock's permanent number in CRSP, blank for the 14 additions whose prices are needed after CRSP_END
# (they are priced from Yahoo Finance, by ticker).
ADDITIONS = """\
name,ticker,announced,included,permno
Bloom Energy,BE,2026-09-04,2026-09-21,
Illumina,ILMN,2026-09-04,2026-09-21,
Everpure,P,2026-09-04,2026-09-21,
Reddit,RDDT,2026-08-13,2026-08-18,
Ferguson Enterprises,FERG,2026-07-31,2026-08-05,
Flex,FLEX,2026-06-05,2026-06-22,
Marvell Technology,MRVL,2026-06-05,2026-06-22,
Veeva Systems,VEEV,2026-04-30,2026-05-07,
Casey's General Stores,CASY,2026-04-06,2026-04-09,
Coherent,COHR,2026-03-06,2026-03-23,
Lumentum Holdings,LITE,2026-03-06,2026-03-23,
EchoStar,ECHO,2026-03-06,2026-03-23,
Vertiv Holdings,VRT,2026-03-06,2026-03-23,
Ciena,CIEN,2026-02-04,2026-02-09,
CRH plc,CRH,2025-12-05,2025-12-22,24294
Carvana,CVNA,2025-12-05,2025-12-22,16655
Comfort Systems USA,FIX,2025-12-05,2025-12-22,85059
Ares Management,ARES,2025-12-08,2025-12-11,14617
Sandisk,SNDK,2025-11-24,2025-11-28,26181
AppLovin,APP,2025-09-05,2025-09-22,20894
Emcor Group,EME,2025-09-05,2025-09-22,82694
Robinhood Markets,HOOD,2025-09-05,2025-09-22,21723
Interactive Brokers Group,IBKR,2025-08-25,2025-08-28,92043
Block,XYZ,2025-07-18,2025-07-23,15826
The Trade Desk,TTD,2025-07-14,2025-07-18,16309
Datadog,DDOG,2025-07-02,2025-07-09,18911
Coinbase Global,COIN,2025-05-12,2025-05-19,20892
DoorDash,DASH,2025-03-07,2025-03-24,20189
Expand Energy,EXE,2025-03-07,2025-03-24,20583
TKO Group Holdings,TKO,2025-03-07,2025-03-24,87356
Williams-Sonoma,WSM,2025-03-07,2025-03-24,83011
Apollo Global Management,APO,2024-12-06,2024-12-23,12591
Lennox International,LII,2024-12-18,2024-12-23,87034
Workday,WDAY,2024-12-06,2024-12-23,13628
Texas Pacific Land,TPL,2024-11-21,2024-11-26,20391
Dell Technologies,DELL,2024-09-06,2024-09-23,18267
Erie Indemnity,ERIE,2024-09-06,2024-09-23,82486
Palantir Technologies,PLTR,2024-09-06,2024-09-23,19788
CrowdStrike Holdings,CRWD,2024-06-07,2024-06-24,18726
GoDaddy,GDDY,2024-06-07,2024-06-24,15315
KKR,KKR,2024-06-07,2024-06-24,11990
Vistra,VST,2024-05-03,2024-05-08,16736
Deckers Outdoor,DECK,2024-03-01,2024-03-18,79686
Super Micro Computer,SMCI,2024-03-01,2024-03-18,91907
Builders FirstSource,BLDR,2023-12-01,2023-12-18,90720
Jabil,JBL,2023-12-01,2023-12-18,79094
Uber Technologies,UBER,2023-12-01,2023-12-18,18576
Hubbell,HUBB,2023-10-13,2023-10-18,32942
Lululemon Athletica,LULU,2023-10-13,2023-10-18,92203
Airbnb,ABNB,2023-09-01,2023-09-18,20190
Blackstone,BX,2023-09-01,2023-09-18,92108
Kenvue,KVUE,2023-08-09,2023-08-25,23876
Palo Alto Networks,PANW,2023-06-02,2023-06-20,13511
Axon Enterprise,AXON,2023-05-01,2023-05-04,89031
Fair Isaac,FICO,2023-03-03,2023-03-20,11533
Bunge,BG,2023-03-13,2023-03-15,89138
Insulet,PODD,2023-03-10,2023-03-15,92050
Steel Dynamics,STLD,2022-12-19,2022-12-22,84262
First Solar,FSLR,2022-12-12,2022-12-19,91611
Arch Capital Group,ACGL,2022-10-27,2022-11-01,82276
Targa Resources,TRGP,2022-10-06,2022-10-12,12476
EQT,EQT,2022-09-23,2022-10-03,24328
PG&E,PCG,2022-09-23,2022-10-03,13688
CoStar Group,CSGP,2022-09-02,2022-09-19,86288
Invitation Homes,INVH,2022-09-02,2022-09-19,16581
Keurig Dr Pepper,KDP,2022-06-03,2022-06-21,17942
ON Semiconductor,ON,2022-06-03,2022-06-21,88182
VICI Properties,VICI,2022-06-03,2022-06-08,17307
Camden Property Trust,CPT,2022-03-29,2022-04-04,79452
Molina Healthcare,MOH,2022-02-25,2022-03-02,89781
Nordson,NDSN,2022-02-10,2022-02-15,57809
FactSet Research Systems,FDS,2021-12-03,2021-12-20,83597
Signature Bank,SBNY,2021-12-03,2021-12-20,90090
SolarEdge Technologies,SEDG,2021-12-03,2021-12-20,15272
EPAM Systems,EPAM,2021-12-07,2021-12-14,13210
Brown,BRO,2021-09-03,2021-09-20,63467
Ceridian HCM,CDAY,2021-09-03,2021-09-20,17700
Match Group,MTCH,2021-09-03,2021-09-20,15850
Bio-Techne,TECH,2021-08-24,2021-08-30,75694
Moderna,MRNA,2021-07-15,2021-07-21,18312
Charles River Laboratories International,CRL,2021-05-10,2021-05-14,88281
PTC,PTC,2021-04-15,2021-04-20,75912
Caesars Entertainment,CZR,2021-03-12,2021-03-22,14882
Generac Holdings,GNRC,2021-03-12,2021-03-22,93246
NXP Semiconductors,NXPI,2021-03-12,2021-03-22,12084
Penn National Gaming,PENN,2021-03-12,2021-03-22,80563
Monolithic Power Systems,MPWR,2021-02-08,2021-02-12,90455
Trimble,TRMB,2021-01-15,2021-01-21,76230
Enphase Energy,ENPH,2020-12-30,2021-01-07,13323
Tesla,TSLA,2020-11-16,2020-12-21,93436
Pool,POOL,2020-10-01,2020-10-07,82515
Catalent,CTLT,2020-09-04,2020-09-21,14763
Etsy,ETSY,2020-09-04,2020-09-21,15291
Teradyne,TER,2020-09-04,2020-09-21,51369
Bio-Rad Laboratories,BIO,2020-06-12,2020-06-22,61516
Teledyne Technologies,TDY,2020-06-12,2020-06-22,87445
Tyler Technologies,TYL,2020-06-12,2020-06-22,76185
West Pharmaceutical Services,WST,2020-05-18,2020-05-22,62498
Domino's Pizza,DPZ,2020-05-06,2020-05-12,90248
DexCom,DXCM,2020-05-06,2020-05-12,90664
Paycom Software,PAYC,2020-01-22,2020-01-28,14579
Live Nation Entertainment,LYV,2019-12-13,2019-12-23,91041
STERIS,STE,2019-12-13,2019-12-23,77649
Zebra Technologies,ZBRA,2019-12-13,2019-12-23,76795
Old Dominion Freight Line,ODFL,2019-12-02,2019-12-09,77037
W.R. Berkley,WRB,2019-11-27,2019-12-05,17743
ServiceNow,NOW,2019-11-19,2019-11-21,13447
Las Vegas Sands,LVS,2019-09-26,2019-10-03,90505
NVR,NVR,2019-09-20,2019-09-26,79785
CDW,CDW,2019-09-17,2019-09-23,13949
IDEX,IEX,2019-08-01,2019-08-09,75591
Leidos Holdings,LDOS,2019-08-01,2019-08-09,91547
T-Mobile US,TMUS,2019-07-09,2019-07-15,91937
MarketAxess Holdings,MKTX,2019-06-24,2019-07-01,90454
Wabtec,WAB,2019-02-21,2019-02-27,81677
Atmos Energy,ATO,2019-02-08,2019-02-15,32986
Teleflex,TFX,2019-01-15,2019-01-18,44329
First Republic Bank,FRC,2018-12-27,2019-01-02,12448
Celanese,CE,2018-12-19,2018-12-24,90520
Diamondback Energy,FANG,2018-11-26,2018-12-03,13641
Lamb Weston Holdings,LW,2018-11-26,2018-12-03,16431
Maxim Integrated Products,MXIM,2018-11-26,2018-12-03,11896
Jack Henry & Associates,JKHY,2018-11-07,2018-11-13,88664
Keysight Technologies,KEYS,2018-10-30,2018-11-06,14939
Fortinet,FTNT,2018-10-04,2018-10-11,93132
Rollins,ROL,2018-09-25,2018-10-01,36003
WellCare Health Plans,WCG,2018-09-11,2018-09-17,90272
Arista Networks,ANET,2018-08-23,2018-08-28,14714
Copart,CPRT,2018-06-25,2018-07-02,80320
FleetCor Technologies,FLT,2018-06-15,2018-06-20,12449
Broadridge Financial Solutions,BR,2018-06-08,2018-06-18,91849
HollyFrontier,HFC,2018-06-08,2018-06-18,32803
Twitter,TWTR,2018-06-04,2018-06-07,14295
Evergy,EVRG,2018-05-31,2018-06-05,24053
ABIOMED,ABMD,2018-05-25,2018-05-31,75107
MSCI,MSCI,2018-03-28,2018-04-04,92402
Nektar Therapeutics,NKTR,2018-03-09,2018-03-19,80539
SVB Financial,SIVB,2018-03-09,2018-03-19,11786
Take-Two Interactive Software,TTWO,2018-03-09,2018-03-19,84761
IPG Photonics,IPGP,2018-03-02,2018-03-07,91668
Huntington Ingalls Industries,HII,2017-12-28,2018-01-03,12623
Norwegian Cruise Line,NCLH,2017-10-04,2017-10-13,13760
Cadence Design Systems,CDNS,2017-09-08,2017-09-18,11403
SBA Communications,SBAC,2017-08-24,2017-09-01,86996
Quintiles IMS Holdings,IQV,2017-08-24,2017-08-29,13911
A.O. Smith,AOS,2017-07-19,2017-07-26,65402
Duke Realty,DRE,2017-07-19,2017-07-26,75341
MGM Resorts,MGM,2017-07-19,2017-07-26,11891
Packaging Corporation of America,PKG,2017-07-19,2017-07-26,87541
ResMed,RMD,2017-07-19,2017-07-26,81736
Align Technology,ALGN,2017-06-09,2017-06-19,88860
ANSYS,ANSS,2017-06-09,2017-06-19,83621
Hilton Worldwide,HLT,2017-06-09,2017-06-19,14338
Everest Re Group,RE,2017-06-12,2017-06-19,82303
IHS Markit,INFO,2017-05-24,2017-06-02,14704
Gartner,IT,2017-03-29,2017-04-05,79698
Advanced Micro Devices,AMD,2017-03-10,2017-03-20,61241
Alexandria Real Estate Equities,ARE,2017-03-10,2017-03-20,84767
Raymond James Financial,RJF,2017-03-10,2017-03-20,69649
Synopsys,SNPS,2017-03-13,2017-03-16,77357
Dish Network,DISH,2017-03-06,2017-03-13,81696
Regency Centers,REG,2017-02-23,2017-03-02,79782
CBOE Holdings,CBOE,2017-02-23,2017-03-01,93429
Incyte,INCY,2017-02-23,2017-02-28,79906
IDEXX Laboratories,IDXX,2017-01-03,2017-01-05,76709
Mid-America Apartment Communities,MAA,2016-11-29,2016-12-02,80183
Coty,COTY,2016-09-27,2016-10-03,13983
The Cooper Companies,COO,2016-09-20,2016-09-23,65541
Charter Communications,CHTR,2016-08-31,2016-09-08,12308
Mettler-Toledo International,MTD,2016-08-25,2016-09-06,85621
Albemarle,ALB,2016-06-23,2016-07-01,80191
Alliant Energy,LNT,2016-06-29,2016-07-01,58819
Fortune Brands Home & Security,FBHS,2016-06-21,2016-06-24,12981
TransDigm Group,TDG,2016-05-27,2016-06-03,91152
Arthur J. Gallagher,AJG,2016-05-24,2016-05-31,38093
LKQ,LKQ,2016-05-18,2016-05-23,89866
Digital Realty Trust,DLR,2016-05-13,2016-05-18,90373
Alaska Air Group,ALK,2016-05-10,2016-05-13,28804
Acuity Brands,AYI,2016-04-26,2016-05-03,89216
Global Payments,GPN,2016-04-19,2016-04-25,88873
"Ulta Salon, Cosmetics & Fragrance",ULTA,2016-04-07,2016-04-18,92322
Foot Locker,FL,2016-03-28,2016-04-04,15456
Centene,CNC,2016-03-24,2016-03-30,89269
Hologic,HOLX,2016-03-24,2016-03-30,76095
UDR,UDR,2016-03-03,2016-03-07,59300
American Water Works,AWK,2016-03-01,2016-03-04,92614
Concho Resources,CXO,2016-02-16,2016-02-22,92239
Citizens Financial Group,CFG,2016-01-26,2016-02-01,14889
Federal Realty Investment Trust,FRT,2016-01-22,2016-02-01,58413
Extra Space Storage,EXR,2016-01-13,2016-01-19,90305
Church & Dwight,CHD,2015-12-22,2015-12-29,23393
Illumina,ILMN,2015-11-12,2015-11-19,88446
Synchrony Financial,SYF,2015-11-09,2015-11-18,14776
Verisk Analytics,VRSK,2015-09-30,2015-10-08,93089
United Continental Holdings,UAL,2015-08-27,2015-09-03,91103
Activision Blizzard,ATVI,2015-08-27,2015-08-31,79678
Signet Jewelers,SIG,2015-07-24,2015-07-29,12076
Advance Auto Parts,AAP,2015-07-06,2015-07-09,89217
J.B. Hunt Transport Services,JBHT,2015-06-26,2015-07-01,42877
Qorvo,QRVO,2015-06-09,2015-06-12,85035
Realty Income,O,2015-03-27,2015-04-07,81045
American Airlines Group,AAL,2015-03-16,2015-03-23,21020
Equinix,EQIX,2015-03-13,2015-03-23,89617
Hanesbrands,HBI,2015-03-13,2015-03-23,91416
SL Green Realty,SLG,2015-03-13,2015-03-23,85265
Henry Schein,HSIC,2015-03-13,2015-03-18,82581
Skyworks Solutions,SWKS,2015-03-09,2015-03-12,45911
Endo International,ENDP,2015-01-21,2015-01-27,88436
HCA Holdings,HCA,2015-01-23,2015-01-27,12622
Royal Caribbean Cruises,RCL,2014-12-01,2014-12-05,79145
Level 3 Communications,LVLT,2014-10-29,2014-11-05,85991
Universal Health Services,UHS,2014-09-12,2014-09-22,79637
United Rentals,URI,2014-09-12,2014-09-22,85663
Mallinckrodt,MNK,2014-08-14,2014-08-19,14011
Martin Marietta Materials,MLM,2014-06-27,2014-07-02,80204
Affiliated Managers Group,AMG,2014-06-24,2014-07-01,85593
Cimarex Energy,XEC,2014-06-12,2014-06-20,89509
Avago Technologies,AVGO,2014-05-05,2014-05-08,93002
Under Armour,UAA,2014-04-24,2014-05-01,90979
Essex Property,ESS,2014-03-26,2014-04-02,80681
Keurig Green Mountain,GMCR,2014-03-14,2014-03-24,79588
Tractor Supply,TSCO,2014-01-16,2014-01-24,80286
Alliance Data Systems,ADS,2013-12-11,2013-12-23,89002
Facebook,FB,2013-12-11,2013-12-23,13407
Mohawk Industries,MHK,2013-12-11,2013-12-23,77496
General Growth Properties,GGP,2013-12-04,2013-12-10,79129
Michael Kors,KORS,2013-10-28,2013-11-13,13142
Transocean,RIG,2013-10-21,2013-10-29,79237
Ametek,AME,2013-09-11,2013-09-23,85257
Vertex,VRTX,2013-09-11,2013-09-23,76744
Delta Air Lines,DAL,2013-09-06,2013-09-11,91926
Nielsen,NLSN,2013-07-01,2013-07-09,12542
Zoetis,ZTS,2013-06-14,2013-06-24,13788
General Motors,GM,2013-06-03,2013-06-07,12369
Kansas City Southern,KSU,2013-05-16,2013-05-24,12650
Macerich,MAC,2013-05-03,2013-05-09,80399
Regeneron,REGN,2013-04-24,2013-05-01,76614
PVH,PVH,2013-02-07,2013-02-14,13936
Delphi Automotive,APTV,2012-12-18,2012-12-24,13103
Garmin,GRMN,2012-12-05,2012-12-12,88837
Dollar General,DG,2012-11-26,2012-12-03,93096
PetSmart,PETM,2012-10-01,2012-10-05,79411
LyondellBasell Industries,LYB,2012-08-29,2012-09-05,12345
Ensco plc,ESV,2012-07-26,2012-07-31,93159
Seagate Technology,STX,2012-06-25,2012-07-02,89641
Monster Beverage,MNST,2012-06-21,2012-06-29,88031
Lam Research,LRCX,2012-05-10,2012-06-05,48486
Alexion Pharmaceuticals,ALXN,2012-05-21,2012-05-25,83111
Kinder Morgan Inc.,KMI,2012-05-17,2012-05-25,12558
Fossil Inc.,FOSL,2012-03-28,2012-04-04,79089
Crown Castle International,CCI,2012-03-07,2012-03-13,86339
BorgWarner,BWA,2011-12-08,2011-12-19,79545
Dollar Tree,DLTR,2011-12-08,2011-12-19,81481
Perrigo,PRGO,2011-12-08,2011-12-19,77182
Cooper Industries plc,CBE,2011-11-18,2011-11-23,21979
TE Connectivity,TEL,2011-10-11,2011-10-17,92157
Mosaic,MOS,2011-09-21,2011-09-26,90386
Accenture plc,ACN,2011-06-27,2011-07-06,89071
Alpha Natural Resources,ANRZQ,2011-05-24,2011-06-02,90547
Chipotle Mexican Grill,CMG,2011-04-20,2011-04-28,91068
BlackRock,BLK,2011-03-30,2011-04-04,87267
Edwards Lifesciences,EW,2011-03-24,2011-04-01,87657
Covidien,COV,2011-01-26,2011-03-01,92156
Joy Global,JOY,2011-02-16,2011-02-28,89105
Noble Corp.,NE,2011-01-12,2011-01-18,90537
Cablevision Systems,CVC,2010-12-09,2010-12-20,68857
F5 Networks,FFIV,2010-12-09,2010-12-20,86964
Netflix,NFLX,2010-12-09,2010-12-20,89393
Newfield Exploration,NFX,2010-12-09,2010-12-20,79915
Ingersoll-Rand plc,IR,2010-11-10,2010-11-17,12431
CarMax,KMX,2010-06-22,2010-06-28,89508
Cerner Corp.,CERN,2010-03-18,2010-04-30,10909
ONEOK Inc.,OKE,2010-03-03,2010-03-15,25232
Discovery Communications (Series A),DISCA,2010-02-16,2010-03-01,90805
Helmerich & Payne,HP,2010-02-22,2010-03-01,32707
Berkshire Hathaway Class B,BRK.B,2010-01-26,2010-02-16,83443
Urban Outfitters,URBN,2010-01-29,2010-02-08,79881
NRG Energy,NRG,2010-01-21,2010-01-29,90071
"""

# Additions effective at S&P's scheduled quarterly rebalance: the Monday after the third Friday of March, June,
# September or December (or the next trading day), announced at least 7 calendar days before. These seven take
# effect on a rebalance date but replace a company that left because of a takeover, spinoff or conversion (checked
# against S&P's press releases), so they are not part of the scheduled batch.
OFF_CYCLE = {
    ("CVC", "2010-12-09"): "replaced King Pharmaceuticals (acquired by Pfizer)",
    ("ZTS", "2013-06-14"): "added because Pfizer was splitting off its Zoetis stake",
    ("AAL", "2015-03-16"): "replaced Allergan (acquired by Actavis)",
    ("HLT", "2017-06-09"): "replaced Yahoo! (converting to a closed-end fund)",
    ("RE", "2017-06-12"): "replaced Mead Johnson (acquired by Reckitt Benckiser)",
    ("CDNS", "2017-09-08"): "replaced Staples (acquired by Sycamore Partners)",
    ("FSLR", "2022-12-12"): "replaced Fortune Brands (moved to MidCap 400 after its MasterBrand spinoff)",
}

# CRSP has two formats: the legacy daily stock file (crsp.dsf), which stops at the end of 2024, and the new one
# (crsp.dsf_v2), which carries on. Each date is taken from the first table that has it. Column names are looked up.
PRICE_TABLES = [
    ("dsf", dict(date=("date",), open=("openprc",), close=("prc",), ret=("ret",), retx=("retx",))),
    ("dsf_v2", dict(date=("dlycaldt", "date"), open=("dlyopen", "openprc"), close=("dlyprc", "prc"),
                    ret=("dlyret", "ret"), retx=("dlyretx", "retx"))),
]


def additions():
    """The list above as a table (in its own order, newest first), with each addition's reference number (E001 =
    the oldest) and whether it took effect at a scheduled quarterly rebalance."""
    ev = pd.read_csv(io.StringIO(ADDITIONS), dtype={"name": str, "ticker": str, "permno": "Int64"})
    ev["announced"] = pd.to_datetime(ev.announced)
    ev["included"] = pd.to_datetime(ev.included)
    n = len(ev)
    order = np.lexsort((np.arange(n), ev.announced.values))          # oldest first; same day: list order
    ref = np.empty(n, dtype=object)
    ref[order] = [f"E{j + 1:03d}" for j in range(n)]
    ev["ref"] = ref

    def third_friday(y, m):
        d = dt.date(y, m, 15)
        return d + dt.timedelta((4 - d.weekday()) % 7)
    inc = [d.date() for d in ev.included]
    on_rebalance = np.array([d.month in (3, 6, 9, 12) and 1 <= (d - third_friday(d.year, d.month)).days <= 4
                             for d in inc])
    notice = ((ev.included - ev.announced).dt.days >= 7).values
    off = np.array([(t, a.strftime("%Y-%m-%d")) in OFF_CYCLE for t, a in zip(ev.ticker, ev.announced)])
    ev["quarterly"] = np.where(on_rebalance & notice & ~off, "Yes", "No")
    return ev


def french_factors():
    """Kenneth French's daily Fama-French factors (Mkt-RF, SMB, HML, RF) and momentum (Mom), as decimals, from
    FIRST_DAY to FF_END. No login needed."""
    def table(zipname):
        req = urllib.request.Request(FRENCH + zipname, headers={"User-Agent": "Mozilla/5.0"})
        try:
            raw = urllib.request.urlopen(req, timeout=60).read()
        except Exception as err:
            raise SystemExit(f"Could not download {zipname} from Kenneth French's data library ({err}). "
                             f"Check the internet connection and run the script again.")
        z = zipfile.ZipFile(io.BytesIO(raw))
        lines = z.read(z.namelist()[0]).decode("latin-1").splitlines()
        header, rows = None, []
        for line in lines:
            parts = [p.strip() for p in line.split(",")]
            if re.fullmatch(r"\d{8}", parts[0]):
                rows.append(parts)
            elif rows:
                break                                    # first table only (daily); stop at the end of it
            elif len(parts) > 1:
                header = parts[1:]                       # the last comma line before the data holds the names
        df = pd.DataFrame(rows, columns=["date"] + header)
        df["date"] = pd.to_datetime(df.date, format="%Y%m%d")
        return df.set_index("date").apply(pd.to_numeric, errors="coerce") / 100   # percent -> decimal

    ff = table("F-F_Research_Data_Factors_daily_CSV.zip").join(table("F-F_Momentum_Factor_daily_CSV.zip"), how="left")
    ff.columns = [c.replace(" ", "") for c in ff.columns]
    if ff.index.max() < FF_END:
        raise SystemExit(f"French's factors end on {ff.index.max().date()}, before {FF_END.date()}.")
    ff = ff[(ff.index >= FIRST_DAY) & (ff.index <= FF_END)]
    print(f"  Fama-French daily factors: {ff.index.min().date()} to {ff.index.max().date()}")
    return ff


def crsp_prices(ev):
    """Daily open, close and returns from CRSP, through WRDS: SPY from FIRST_DAY, and each addition with a PERMNO
    from HIST_DAYS trading days before its announcement to its inclusion day. Nothing after CRSP_END is used.
    Returns SPY (indexed by date) and the additions' rows (one per stock and day)."""
    try:
        import wrds
    except ImportError:
        raise SystemExit("This script needs the wrds and yfinance packages. In Colab run:  !pip install wrds yfinance")
    print("  Connecting to WRDS: type your WRDS username and password when asked, then approve the Duo push.")
    db = wrds.Connection()
    try:
        sources = []
        for table, cands in PRICE_TABLES:
            try:
                have = set(db.describe_table("crsp", table)["name"].str.lower())
            except Exception as err:
                print(f"  crsp.{table}: not available ({err})")
                continue
            sel = {k: next((c for c in v if c in have), None) for k, v in cands.items()}
            if not (sel["date"] and sel["open"] and sel["close"]):
                print(f"  crsp.{table}: columns not recognised, skipped")
                continue
            last = pd.Timestamp(db.raw_sql(f"select max({sel['date']}) as d from crsp.{table} "
                                           f"where permno = {SPY_PERMNO}").d.iloc[0])
            sources.append(dict(table=table, sel=sel, last=min(last, CRSP_END)))
            print(f"  crsp.{table}: data to {last.date()}, used to {min(last, CRSP_END).date()}")
        if not sources or max(s["last"] for s in sources) < CRSP_END:
            raise SystemExit(f"CRSP's daily stock file does not reach {CRSP_END.date()} on this WRDS account.")

        def pull_all(where):
            """Rows meeting where(date column), each date from the first table that has it."""
            out, covered = [], pd.Timestamp("1900-01-01")
            for s in sources:
                if s["last"] <= covered:
                    continue
                sel, dc = s["sel"], s["sel"]["date"]
                fields = ", ".join(f"{c} as {k}" for k, c in sel.items() if c)
                out.append(db.raw_sql(f"select permno, {fields} from crsp.{s['table']} where ({where(dc)}) "
                                      f"and {dc} > '{covered.date()}' and {dc} <= '{s['last'].date()}'",
                                      date_cols=["date"]))
                covered = s["last"]
            df = pd.concat([o for o in out if len(o)] or out, ignore_index=True)
            for k in ["open", "close", "ret", "retx"]:
                df[k] = pd.to_numeric(df[k], errors="coerce").astype(float) if k in df else np.nan
            df["close"] = df.close.abs()                 # CRSP: a negative price is a bid/ask midpoint (no trade)
            return df

        spy = pull_all(lambda dc: f"permno = {SPY_PERMNO} and {dc} >= '{FIRST_DAY.date()}'")
        spy = spy.drop_duplicates("date").sort_values("date").set_index("date")[["open", "close"]]
        cal = spy.index
        crsp_ev = ev[ev.permno.notna()]
        start = {e.Index: cal[max(cal.searchsorted(e.announced) - HIST_DAYS, 0)] for e in crsp_ev.itertuples()}
        parts = []
        for i in range(0, len(crsp_ev), 60):                # 60 additions per query
            chunk = crsp_ev.iloc[i:i + 60]
            parts.append(pull_all(lambda dc, chunk=chunk: " or ".join(
                f"(permno = {int(e.permno)} and {dc} between '{start[e.Index].date()}' and '{e.included.date()}')"
                for e in chunk.itertuples())))
            print(f"  CRSP prices: {min(i + 60, len(crsp_ev))} of {len(crsp_ev)} additions")
    finally:
        db.close()
    stocks = pd.concat(parts, ignore_index=True).drop_duplicates(["permno", "date"])
    rows = []
    for e in crsp_ev.itertuples():
        d = stocks[(stocks.permno == e.permno) & stocks.date.between(start[e.Index], e.included)]
        rows.append(d.assign(event=e.Index, src="crsp"))
    return spy, pd.concat(rows, ignore_index=True)


def yahoo_prices(ticker, start, end):
    """Daily open and close from Yahoo Finance (split-adjusted, not dividend-adjusted), with ret from Adj Close
    (dividends included) and retx from Close. Starts 15 days early so the first day has a return."""
    try:
        import yfinance as yf
    except ImportError:
        raise SystemExit("This script needs the wrds and yfinance packages. In Colab run:  !pip install wrds yfinance")
    h = None
    for attempt in range(4):
        try:
            h = yf.Ticker(ticker).history(start=start - pd.Timedelta(days=15), end=end + pd.Timedelta(days=1),
                                          auto_adjust=False, actions=True)
            if h is not None and len(h):
                break
        except Exception as err:
            print(f"  Yahoo Finance, {ticker}: {err}; trying again")
        if attempt < 3:
            time.sleep(15 * 2 ** attempt)                # 15, 30 and 60 seconds
    if h is None or not len(h):
        raise SystemExit(f"Yahoo Finance returned no prices for {ticker}. Run the script again in a few minutes (if it "
                         f"keeps failing, run !pip install -U yfinance in a cell, restart the session and try again).")
    idx = pd.to_datetime(h.index)
    h.index = (idx.tz_localize(None) if idx.tz is not None else idx).normalize()
    out = pd.DataFrame({"open": h["Open"], "close": h["Close"], "ret": h["Adj Close"].pct_change(),
                        "retx": h["Close"].pct_change()})
    out.index.name = "date"
    return out[(out.index >= start) & (out.index <= end)]


def download_data():
    """PART 1: download everything and lay it out for the calculations (build_inputs)."""
    print("PART 1. Downloading the data")
    ev = additions()
    ff = french_factors()
    # Yahoo Finance first (no login needed, so a Yahoo failure does not waste a WRDS login): SPY after CRSP_END and the
    # additions without a PERMNO (420 calendar days back is always more than HIST_DAYS trading days)
    ys = yahoo_prices("SPY", CRSP_END - pd.Timedelta(days=40), pd.Timestamp.today().normalize())
    yh = {e.Index: yahoo_prices(e.ticker.replace(".", "-"), e.announced - pd.Timedelta(days=420), e.included)
          for e in ev[ev.permno.isna()].itertuples()}
    print(f"  Yahoo Finance: SPY to {ys.index.max().date()} and {len(yh)} additions")
    spy, daily = crsp_prices(ev)
    # SPY: CRSP to CRSP_END, then Yahoo Finance (with a check on the days both sources have)
    seam = ys.index.intersection(spy.index)
    if len(seam):
        gap = (ys.loc[seam, "close"] / spy.loc[seam, "close"] - 1).abs().max()
        print(f"  SPY: Yahoo Finance and CRSP closes differ by at most {gap:.3%} on {len(seam)} shared days")
    spy = pd.concat([spy, ys.loc[ys.index > CRSP_END, ["open", "close"]]]).sort_index()
    cal = spy.index                                  # the trading-day calendar: the days SPY traded
    # the additions without a PERMNO, from Yahoo Finance by ticker
    rows = [daily]
    for e in ev[ev.permno.isna()].itertuples():
        start = cal[max(cal.searchsorted(e.announced) - HIST_DAYS, 0)]
        rows.append(yh[e.Index].loc[start:].reset_index().assign(event=e.Index, src="yahoo"))
    daily = pd.concat(rows, ignore_index=True)
    daily["date"] = pd.to_datetime(daily["date"])
    return build_inputs(ev, spy, daily, ff)


def build_inputs(ev, spy, daily, ff):
    """Lay the downloaded data out the way the calculations read them (the workbook's data sheets):

    ev    one row per addition: dates, quarterly flag and the prices of the stock and of SPY at the announcement
          close, the Ann+1 open (the buy), the close the day before inclusion and the inclusion open and close
    win   (additions x days) arrays of date, open and close for every trading day from the announcement (day 0)
          to the inclusion day, for the stock and for SPY (the Event Window sheets)
    hist  (additions x 261) arrays: column k = k trading days before the announcement, k = 1..260 (the 260-day
          sheets): the stock's close, SPY's date and close
    ff    the Fama-French factors, one row per day

    Stock prices are put on the share basis of the announcement day: a split shows up as a day where the change
    in CRSP's price disagrees with CRSP's price return (retx), and the prices before it are rescaled. So on and after
    the announcement day the prices are CRSP's own. Yahoo's prices are already split-adjusted."""
    cal = spy.index
    n = len(ev)
    i_ann, i_inc = cal.searchsorted(ev.announced.values), cal.searchsorted(ev.included.values)
    span = i_inc - i_ann                              # trading days from the announcement to the inclusion day
    adj, splits, dividends = {}, 0, []
    for i, d in daily.groupby("event"):
        d = d.sort_values("date").copy()
        d["rel_day"] = cal.searchsorted(d.date) - i_ann[i]
        ratio = (d.close / d.close.shift(1)) / (1 + d.retx)
        ratio = ratio.where(np.log(ratio.abs()).abs() > 0.01, 1.0).fillna(1.0)
        ratio[d.src == "yahoo"] = 1.0
        f = ratio.cumprod()
        base = f[d.rel_day <= 0].iloc[-1] if (d.rel_day <= 0).any() else f.iloc[0]
        k = f / base
        d["open"], d["close"] = d.open / k, d.close / k
        held = (d.rel_day >= 0) & (d.rel_day <= span[i])
        splits += bool((ratio[held] != 1).any())
        hold = d[(d.rel_day > 1) & (d.rel_day <= span[i])]           # dividends paid after the Ann+1 open
        dividends.append(float((hold.ret - hold.retx).fillna(0).sum()))
        adj[i] = d.drop_duplicates("rel_day").set_index("rel_day")

    nwin = int(span.max()) + 1
    nat = np.datetime64("NaT", "D")
    win = SimpleNamespace(date=np.full((n, nwin), nat), open=np.full((n, nwin), np.nan),
                          close=np.full((n, nwin), np.nan), spy_date=np.full((n, nwin), nat),
                          spy_open=np.full((n, nwin), np.nan), spy_close=np.full((n, nwin), np.nan))
    hist = SimpleNamespace(close=np.full((n, HIST_DAYS + 1), np.nan), spy_date=np.full((n, HIST_DAYS + 1), nat),
                           spy_close=np.full((n, HIST_DAYS + 1), np.nan))
    spy_open, spy_close = spy.open.values, spy.close.values
    day64 = lambda ts: np.datetime64(pd.Timestamp(ts).date(), "D")
    for i in range(n):
        d = adj.get(i)
        for k in range(-HIST_DAYS, span[i] + 1):
            j = i_ann[i] + k
            have = d is not None and k in d.index
            if k >= 0:
                if have:
                    win.date[i, k], win.open[i, k], win.close[i, k] = day64(d.at[k, "date"]), d.at[k, "open"], \
                        d.at[k, "close"]
                if 0 <= j < len(cal):
                    win.spy_date[i, k], win.spy_open[i, k], win.spy_close[i, k] = day64(cal[j]), spy_open[j], \
                        spy_close[j]
            else:
                if have:
                    hist.close[i, -k] = d.at[k, "close"]
                if 0 <= j < len(cal):
                    hist.spy_date[i, -k], hist.spy_close[i, -k] = day64(cal[j]), spy_close[j]

    gaps = [ev.name.iat[i] for i in range(n) if (win.date[i, :span[i] + 1] != win.spy_date[i, :span[i] + 1]).any()
            or np.isnan(win.close[i, :span[i] + 1]).any() or np.isnan(win.spy_close[i, :span[i] + 1]).any()]
    if gaps:
        print(f"  Warning: a trading day is missing inside the holding period of {', '.join(gaps)}, so the portfolio "
              f"results can differ slightly from the paper")

    rows = np.arange(n)
    E = span
    out = pd.DataFrame({
        "row": rows + 1, "ref": ev.ref.values, "name": ev.name.values, "ticker": ev.ticker.values,
        "announced": pd.to_datetime(ev.announced.values), "included": pd.to_datetime(ev.included.values),
        "quarterly": ev.quarterly.values,
        "s_ann_close": win.close[rows, 0], "s_buy_open": win.open[rows, 1], "s_pre_close": win.close[rows, E - 1],
        "s_inc_open": win.open[rows, E], "s_inc_close": win.close[rows, E],
        "m_ann_close": win.spy_close[rows, 0], "m_buy_open": win.spy_open[rows, 1],
        "m_pre_close": win.spy_close[rows, E - 1], "m_inc_open": win.spy_open[rows, E],
        "m_inc_close": win.spy_close[rows, E]})
    missing = out.columns[out.isna().any()].tolist()
    if missing:
        bad = out[out.isna().any(axis=1)]
        raise SystemExit(f"Prices are missing for {', '.join(bad.name.astype(str))} ({', '.join(missing)}).")
    ffd = pd.DataFrame([{"date": np.datetime64(d.date(), "D"), "mkt": r["Mkt-RF"], "smb": r["SMB"], "hml": r["HML"],
                         "rf": r["RF"], "mom": r["Mom"]} for d, r in ff.iterrows()])
    dividends = np.array(dividends)
    is_crsp, cut = ev.permno.notna().values, np.datetime64(CRSP_END.date(), "D")
    total = lambda *arrays: math.fsum(x for a in arrays for x in np.ravel(a) if not np.isnan(x))
    sums = {"CRSP stock prices": total(hist.close[is_crsp], win.open[is_crsp], win.close[is_crsp]),
            "Yahoo stock prices": total(hist.close[~is_crsp], win.open[~is_crsp], win.close[~is_crsp]),
            "CRSP SPY prices": total(hist.spy_close[hist.spy_date <= cut], win.spy_open[win.spy_date <= cut],
                                     win.spy_close[win.spy_date <= cut]),
            "Yahoo SPY prices": total(hist.spy_close[hist.spy_date > cut], win.spy_open[win.spy_date > cut],
                                      win.spy_close[win.spy_date > cut]),
            "Fama-French factors": total(ffd[["mkt", "smb", "hml", "rf", "mom"]].to_numpy())}
    info = SimpleNamespace(crsp=int(ev.permno.notna().sum()), yahoo=int(ev.permno.isna().sum()),
                           spy_last=cal.max(), splits=splits, n_div=int((np.abs(dividends) > 1e-6).sum()),
                           max_div=float(dividends.max()),
                           changed=[k for k, v in sums.items() if not math.isclose(v, PAPER_DATA[k], rel_tol=1e-10)])
    return SimpleNamespace(n=n, ev=out, win=win, hist=hist, ff=ffd, info=info)


# =====================================================================================================================
# PART 2. THE CALCULATIONS: every statistic of the paper, recomputed from the prices (as in the workbook's sheets)
# =====================================================================================================================

# ---------------------------------------------------------------------------------------------------------------------
# 2.0 Small helpers that copy Excel's functions exactly
# ---------------------------------------------------------------------------------------------------------------------
def to_date(d):
    """numpy/pandas date -> datetime.date (for printing and comparing)."""
    return pd.Timestamp(d).date()


def t_dist_2t(x, df):
    """T.DIST.2T(x, df). Excel and LibreOffice truncate df to an integer (matters only for the Welch test)."""
    return 2.0 * stats.t.sf(abs(x), math.floor(df))


def t_inv_2t(prob, df):
    """T.INV.2T(prob, df): the t value that leaves prob/2 in each tail."""
    return stats.t.ppf(1.0 - prob / 2.0, math.floor(df))


def norm_2t(z):
    """2 * (1 - NORM.S.DIST(|z|, TRUE)), written exactly as on the sheets (it is exactly 0 for |z| > 8.3)."""
    return 2.0 * (1.0 - stats.norm.cdf(abs(z)))


def rank_avg(x):
    """RANK.AVG(value, range, 1): ascending ranks, ties get the average rank."""
    return stats.rankdata(x, method="average")


def stdev_s(x):
    """STDEV.S: sample standard deviation (divides by n - 1)."""
    return float(np.std(np.asarray(x, float), ddof=1))


def text_pct(x, decimals=0):
    """Excel TEXT(x, "0%") / TEXT(x, "0.0%"), used in a few labels."""
    return f"{x * 100:.{decimals}f}%"


def text_dmy(d):
    """Excel TEXT(date, "d mmm yyyy"), e.g. 29 Dec 2017."""
    d = to_date(d)
    return f"{d.day} {d:%b %Y}"



# ---------------------------------------------------------------------------------------------------------------------
# 2.1 SPY's close on each day
# ---------------------------------------------------------------------------------------------------------------------
def spy_close_by_date(raw):
    """SPY's close for every date in the SPY price blocks (hist and win: the Daily Account takes each day's close
    from them). Checks that no date has two different closes."""
    dates = np.concatenate([raw.hist.spy_date[:, 1:].ravel(), raw.win.spy_date.ravel()])
    closes = np.concatenate([raw.hist.spy_close[:, 1:].ravel(), raw.win.spy_close.ravel()])
    ok = ~np.isnat(dates) & ~np.isnan(closes)
    s = pd.Series(closes[ok], index=pd.to_datetime(dates[ok]))
    conflicts = s.groupby(level=0).nunique()
    if (conflicts > 1).any():
        sys.exit(f"SPY sheets disagree on the close for {list(conflicts[conflicts > 1].index[:5])}")
    return s.groupby(level=0).first().sort_index()


# ---------------------------------------------------------------------------------------------------------------------
# 2.2 Trade Returns and Beta Estimation sheets
# ---------------------------------------------------------------------------------------------------------------------
def beta_estimation(raw):
    """Beta Estimation sheet: for each stock, the OLS slope of its daily returns on SPY's over days -250 to -30
    before the announcement. Return day -k = close(day -k) / close(day -k-1) - 1 when both closes exist. A stock
    with fewer than MIN_BETA_DAYS returns gets beta = 1. SLOPE only uses days where both returns exist."""
    k = np.arange(BETA_FIRST_DAY, BETA_LAST_DAY - 1, -1)                  # 250, 249, ..., 30
    stock = raw.hist.close[:, k] / raw.hist.close[:, k + 1] - 1.0
    spy = raw.hist.spy_close[:, k] / raw.hist.spy_close[:, k + 1] - 1.0
    days = (~np.isnan(stock)).sum(axis=1)                                  # COUNT of stock returns (column D)
    slope = np.full(raw.n, np.nan)
    for i in range(raw.n):
        both = ~np.isnan(stock[i]) & ~np.isnan(spy[i])
        x, y = spy[i, both], stock[i, both]
        if len(x) > 1:
            slope[i] = ((x - x.mean()) * (y - y.mean())).sum() / ((x - x.mean()) ** 2).sum()
    used = np.where(days >= MIN_BETA_DAYS, slope, 1.0)                     # column G
    return pd.DataFrame({"days": days, "slope": np.where(days >= MIN_BETA_DAYS, slope, np.nan), "beta": used})


def trade_returns(raw, beta):
    """Trade Returns sheet: one row per trade (workbook column letters in the comments)."""
    ev = raw.ev
    tr = ev[["row", "ref", "name", "ticker", "announced", "included", "quarterly"]].copy()
    tr["year"] = tr.announced.dt.year                                                        # G
    s_ann1_close, m_ann1_close = raw.win.close[:, 1], raw.win.spy_close[:, 1]                # K, Q (Ann+1 close)
    tr["R_A"] = ev.s_inc_open / ev.s_buy_open - 1                                            # U stock, Window A
    tr["M_A"] = ev.m_inc_open / ev.m_buy_open - 1                                            # V SPY, Window A
    tr["AR_A"] = tr.R_A - tr.M_A                                                             # W
    tr["R_B"] = ev.s_inc_close / ev.s_buy_open - 1                                           # X stock, Window B
    tr["M_B"] = ev.m_inc_close / ev.m_buy_open - 1                                           # Y SPY, Window B
    tr["AR_B"] = tr.R_B - tr.M_B                                                             # Z
    tr["beta"] = beta.beta.values                                                            # AA
    tr["AR_A_adj"] = tr.R_A - tr.beta * tr.M_A                                               # AB risk-adjusted A
    tr["AR_B_adj"] = tr.R_B - tr.beta * tr.M_B                                               # AC risk-adjusted B
    tr["V5_stock"] = ev.s_pre_close / ev.s_buy_open - 1                                      # AD V5: sell the day
    tr["V5_spy"] = ev.m_pre_close / ev.m_buy_open - 1                                        # AE before inclusion
    tr["V5"] = tr.V5_stock - tr.V5_spy                                                       # AF
    # Figure 1: abnormal return since the announcement close, at each moment (AG..AL)
    tr["CAR_ann_close"] = ev.s_ann_close / ev.s_ann_close - ev.m_ann_close / ev.m_ann_close  # AG (= 0)
    tr["CAR_buy_open"] = ev.s_buy_open / ev.s_ann_close - ev.m_buy_open / ev.m_ann_close     # AH overnight jump
    tr["CAR_ann1_close"] = s_ann1_close / ev.s_ann_close - m_ann1_close / ev.m_ann_close     # AI
    tr["CAR_pre_close"] = ev.s_pre_close / ev.s_ann_close - ev.m_pre_close / ev.m_ann_close  # AJ
    tr["CAR_inc_open"] = ev.s_inc_open / ev.s_ann_close - ev.m_inc_open / ev.m_ann_close     # AK
    tr["CAR_inc_close"] = ev.s_inc_close / ev.s_ann_close - ev.m_inc_close / ev.m_ann_close  # AL
    # AM holding days: the Event Window block whose date is the inclusion date
    hit = raw.win.date == ev.included.values.astype("datetime64[D]")[:, None]
    if not hit.any(axis=1).all():
        sys.exit("Some inclusion dates are not in the Event Window sheet")
    tr["days"] = hit.argmax(axis=1)
    tr["label"] = [f"{nm} ({y})" for nm, y in zip(tr.name, tr.year)]                        # AR
    # The prices (I..T), also used by the worked examples of Tests 1 and 2
    for c in ["s_ann_close", "s_buy_open", "s_pre_close", "s_inc_open", "s_inc_close",
              "m_ann_close", "m_buy_open", "m_pre_close", "m_inc_open", "m_inc_close"]:
        tr[c] = ev[c].values
    tr["s_ann1_close"], tr["m_ann1_close"] = s_ann1_close, m_ann1_close
    return tr


# ---------------------------------------------------------------------------------------------------------------------
# 2.3 Holding Days and Daily Account sheets
# ---------------------------------------------------------------------------------------------------------------------
def holding_days(raw, tr):
    """Holding Days sheet: one row per position per trading day held (day 1 = Ann+1, day E = inclusion day),
    positions in sheet order bottom-up (oldest first), with each day's piece of the return.

    Window A stock: day 1 open->close, day E previous close->open (sold at the open), other days close->close.
    Window B stock: day 1 open->close, later days close->close (sold at the close). SPY pieces the same way.
    Long + SPY account, Window A: day 1 = SPY previous close->open, then stock open->close (the account sits in SPY
    until it buys); day E = stock previous close->open, then SPY open->close; other days stock close->close.
    Long + SPY account, Window B: day 1 as in A, later days stock close->close.
    Also adds to tr the chain checks of Trade Returns AN..AQ.
    """
    w = raw.win
    parts = []
    for i in reversed(range(raw.n)):
        E = int(tr.days.iat[i])
        d = np.arange(1, E + 1)
        so, sc, sp = w.open[i, 1:E + 1], w.close[i, 1:E + 1], w.close[i, 0:E]           # stock open, close, prev close
        mo, mc, mp = w.spy_open[i, 1:E + 1], w.spy_close[i, 1:E + 1], w.spy_close[i, 0:E]
        first, last = d == 1, (d == E) & (d != 1)                                         # day 1 is tested first
        stock_A = np.where(first, sc / so - 1, np.where(last, so / sp - 1, sc / sp - 1))
        spy_A = np.where(first, mc / mo - 1, np.where(last, mo / mp - 1, mc / mp - 1))
        stock_B = np.where(first, sc / so - 1, sc / sp - 1)
        spy_B = np.where(first, mc / mo - 1, mc / mp - 1)
        long_A = np.where(first, (mo / mp) * (sc / so) - 1, np.where(last, (so / sp) * (mc / mo) - 1, sc / sp - 1))
        long_B = np.where(first, (mo / mp) * (sc / so) - 1, sc / sp - 1)
        parts.append(pd.DataFrame({
            "trade": i, "row": tr.row.iat[i], "name": tr.name.iat[i], "day": d, "date": w.date[i, 1:E + 1],
            "stock_open": so, "stock_close": sc, "stock_prev_close": sp,
            "spy_open": mo, "spy_close": mc, "spy_prev_close": mp,
            "stock_A": stock_A, "spy_A": spy_A, "ar_A": stock_A - spy_A,
            "stock_B": stock_B, "spy_B": spy_B, "ar_B": stock_B - spy_B,
            "long_spy_A": long_A, "long_spy_B": long_B}))
    hd = pd.concat(parts, ignore_index=True)
    hd["date"] = pd.to_datetime(hd.date)
    # Trade Returns AN..AQ: chaining a trade's daily pieces must give back its window return
    chain = hd.groupby("trade").agg(
        sA=("stock_A", lambda s: np.prod(1 + s.values) - 1), sB=("stock_B", lambda s: np.prod(1 + s.values) - 1),
        mA=("spy_A", lambda s: np.prod(1 + s.values) - 1), mB=("spy_B", lambda s: np.prod(1 + s.values) - 1))
    tr["gap_stock_A"] = chain.sA.values - tr.R_A.values                                       # AN
    tr["gap_stock_B"] = chain.sB.values - tr.R_B.values                                       # AO
    tr["gap_spy_A"] = chain.mA.values - tr.M_A.values                                         # AP
    tr["gap_spy_B"] = chain.mB.values - tr.M_B.values                                         # AQ
    return hd


def daily_account(raw, hd, spy):
    """Daily Account sheet: one row per trading day from the first Fama-French day of 2010 to the last inclusion.

    B1 = SPY bought and held. Long + SPY (Window A or B) = on days with open positions, the equal-weighted average
    of the positions' pieces that day (Holding Days); on days with none, SPY (so it equals B1 that day).
    Excess return = return - RF. Value of $1 = cumulative product; drawdown = value / running maximum - 1.
    """
    ff = raw.ff
    has_f = ~np.isnan(ff.mkt.values)
    ff_last = ff.date.values[int(has_f.sum()) - 1]          # FF_LAST: the last day with Mkt-RF
    last_inclusion = raw.ev.included.max().to_datetime64().astype("datetime64[D]")
    dates = ff.date.values[ff.date.values >= np.datetime64(ACCOUNT_START)]
    if dates.max() < last_inclusion:
        # French's factors stop before the last inclusion (FF_END): the account then continues on SPY's trading
        # days up to the last inclusion day, with RF carried forward (as the workbook does).
        extra = spy.index.values.astype("datetime64[D]")
        dates = np.r_[dates, extra[(extra > dates.max()) & (extra <= last_inclusion)]]
    rf_by_date = pd.Series(ff.rf.values, index=pd.to_datetime(ff.date.values))
    rf_last = rf_by_date[pd.Timestamp(ff_last)]
    da = pd.DataFrame({"date": pd.to_datetime(dates)})
    da["rf"] = [rf_by_date.get(d, np.nan) if d <= pd.Timestamp(ff_last) else rf_last for d in da.date]  # C
    held = hd.groupby("date")
    da["positions"] = held.size().reindex(da.date).fillna(0).astype(int).values                       # D
    da["spy_close"] = spy.reindex(da.date).values                                                      # E
    da["price_found"] = ~np.isnan(da.spy_close.values)                                                 # H
    first_prev = spy[spy.index < da.date.iat[0]].iloc[-1]                                              # 31 Dec 2009
    da["spy_prev_close"] = np.r_[first_prev, da.spy_close.values[:-1]]                                 # F
    da["B1"] = da.spy_close / da.spy_prev_close - 1                                                    # G
    da["avg_A"] = held.long_spy_A.mean().reindex(da.date).values                                       # I
    da["avg_B"] = held.long_spy_B.mean().reindex(da.date).values                                       # K
    on = da.positions.values > 0
    da["A"] = np.where(on, da.avg_A.values, da.B1.values)                                              # J
    da["B"] = np.where(on, da.avg_B.values, da.B1.values)                                              # L
    for k in ["A", "B1", "B"]:
        da["x_" + k] = da[k] - da.rf                                                                   # M, N, O
        da["value_" + k] = np.cumprod(1 + da[k].values)                                                # P, Q, R
        v = da["value_" + k].values
        da["dd_" + k] = v / np.maximum.accumulate(v) - 1                                               # S, T, U
    # Fama-French factors on each day (blank after French's last day)
    f = ff.set_index(pd.to_datetime(ff.date.values))
    for c in ["mkt", "smb", "hml", "mom"]:
        da[c] = f[c].reindex(da.date).values
    da.attrs["ff_last"] = pd.Timestamp(ff_last)
    return da


def day_number(dates, d):
    """MATCH(d, dates, 1): 1-based position of the last date on or before d."""
    return int(np.searchsorted(dates.values.astype("datetime64[D]"), np.datetime64(d, "D"), side="right"))


# ---------------------------------------------------------------------------------------------------------------------
# 2.4 Regression Workings: OLS with Newey-West standard errors
# ---------------------------------------------------------------------------------------------------------------------
def newey_west_ols(y, X, lags=NW_LAGS):
    """b = (X'X)^-1 X'y; Newey-West covariance with Bartlett weights 1 - l/(lags+1), no small-sample scaling:
    V = (X'X)^-1 [G0 + sum_l w_l (G_l + G_l')] (X'X)^-1 with G_l = sum_t u_t u_{t-l}', u_t = e_t x_t."""
    y, X = np.asarray(y, float), np.asarray(X, float)
    xtx_inv = np.linalg.inv(X.T @ X)
    b = xtx_inv @ (X.T @ y)
    u = X * (y - X @ b)[:, None]
    S = u.T @ u
    for lag in range(1, lags + 1):
        G = u[lag:].T @ u[:-lag]
        S += max(0.0, 1 - lag / (lags + 1)) * (G + G.T)
    se = np.sqrt(np.diag(xtx_inv @ S @ xtx_inv))
    return b, se


def regression(da, ycol, first_day, last_day, four_factor=False):
    """Regress the daily excess return ycol on Mkt-RF (CAPM) or Mkt-RF, SMB, HML, Mom (four-factor) over day
    numbers first_day..last_day (1-based, inclusive). Returns the coefficients, Newey-West SEs and T."""
    s = da.iloc[first_day - 1:last_day]
    cols = ["mkt", "smb", "hml", "mom"] if four_factor else ["mkt"]
    X = np.column_stack([np.ones(len(s))] + [s[c].values for c in cols])
    b, se = newey_west_ols(s[ycol].values, X)
    return SimpleNamespace(b=b, se=se, T=len(s), alpha_year=b[0] * DAYS_PER_YEAR, t=b[0] / se[0])


# ---------------------------------------------------------------------------------------------------------------------
# 2.5 The statistics used on the Test sheets (as laid out in Test 1-2 Workings and Test 4 Workings)
# ---------------------------------------------------------------------------------------------------------------------
def stat_block(x, dates):
    """Every statistic of Tests 1, 2 and 4 for one series of trade returns x with announcement dates.

    Clustered by date: each announcement date counts once (the average of its trades), G = number of dates.
    Wilcoxon signed rank: exact zeros dropped, average ranks of |x|, normal approximation without tie or
    continuity correction. Sign test: exact binomial, two-sided.
    """
    x = np.asarray(x, float)
    dates = np.asarray(dates)
    N = len(x)
    mean, s = x.mean(), stdev_s(x)
    se = s / math.sqrt(N)
    t = mean / se
    groups = pd.Series(x).groupby(dates)
    by_date = groups.mean()                                    # one value per announcement date
    G = len(by_date)
    meanG, sdG = by_date.mean(), stdev_s(by_date.values)
    seG = sdG / math.sqrt(G)
    tG = meanG / seG
    t_star = t_inv_2t(ALPHA, G - 1)
    nz = x[x != 0]
    n = len(nz)
    ranks = rank_avg(np.abs(nz))
    W = ranks[nz > 0].sum()
    EW, VW = n * (n + 1) / 4, n * (n + 1) * (2 * n + 1) / 24
    z = (W - EW) / math.sqrt(VW)
    k = int((x > 0).sum())
    return SimpleNamespace(
        N=N, mean=mean, s=s, se=se, t=t, p=t_dist_2t(t, N - 1),
        G=G, multi=int((groups.size() > 1).sum()), meanG=meanG, sdG=sdG, seG=seG, tG=tG, pG=t_dist_2t(tG, G - 1),
        t_star=t_star, lo=meanG - t_star * seG, hi=meanG + t_star * seG,
        median=float(np.median(x)), n=n, W=W, EW=EW, VW=VW, z=z, pW=norm_2t(z),
        k=k, beat=k / n, pS=min(1.0, 2 * stats.binom.cdf(min(k, n - k), n, 0.5)),
        mde=MDE_MULT * seG)


def sig_word(p, m):
    """'Significant at the 5% level?' cells."""
    return ("Yes, above zero" if m > 0 else "Yes, below zero") if p < ALPHA else "No"


def verdict(b):
    """Test 1's rule (also used in Tests 2 and 4 and Extra Checks E3)."""
    if b.meanG > 0 and b.pG < ALPHA and b.lo > 0 and b.median > 0 and b.pW < ALPHA and b.z > 0:
        return "Works"
    if b.meanG > 0 and b.pG < ALPHA:
        return "Mixed"
    if (b.meanG <= 0 and b.median <= 0) or (b.meanG < 0 and b.pG < ALPHA) or (b.pG >= ALPHA and b.mde <= MIN_EFFECT):
        return "Doesn't work"
    return "Inconclusive"


def combine(a, b):
    """Tests 1 and 2: one answer from the two windows."""
    if "Works" in (a, b):
        return "Works"
    if a == b == "Doesn't work":
        return "Doesn't work"
    if "Mixed" in (a, b):
        return "Mixed"
    return "Inconclusive"


def yes(cond):
    return "Yes" if cond else "No"


def sharpe_daily(x):
    x = np.asarray(x, float)
    return x.mean() / stdev_s(x)


def jobson_korkie(x1, x2):
    """Jobson-Korkie test with Memmel's (2003) correction, on daily Sharpe ratios of excess returns x1 and x2."""
    sr1, sr2 = sharpe_daily(x1), sharpe_daily(x2)
    rho = np.corrcoef(x1, x2)[0, 1]
    T = len(x1)
    se = math.sqrt((2 - 2 * rho + 0.5 * (sr1 ** 2 + sr2 ** 2 - 2 * sr1 * sr2 * rho ** 2)) / T)
    z = (sr1 - sr2) / se
    return SimpleNamespace(rho=rho, se=se, z=z, p=norm_2t(z))


def drawdown(da, k):
    """Maximum drawdown, the peak before it (first day the pre-trough maximum was reached), the trough, and the
    first later day the value got back to that peak ('Not yet' if never)."""
    v, dd, dates = da["value_" + k].values, da["dd_" + k].values, da.date
    t = int(np.argmin(dd))                              # first day of the lowest drawdown
    peak_value = v[:t + 1].max()
    p = int(np.argmax(v == peak_value))
    after = np.nonzero(v[t + 1:] >= peak_value)[0]
    recovered = to_date(dates.iat[t + 1 + after[0]]) if len(after) else "Not yet"
    return SimpleNamespace(max_dd=dd.min(), peak=to_date(dates.iat[p]), trough=to_date(dates.iat[t]),
                           recovered=recovered)


# ---------------------------------------------------------------------------------------------------------------------
# 2.6 A results sheet: kept cell by cell (Part 3 picks the paper's numbers from it), printed with --details
# ---------------------------------------------------------------------------------------------------------------------
class Report:
    """One results sheet of the workbook. Each row() call records the values under their workbook cell addresses
    (Part 3 reads the paper's numbers from them) and adds a line to the table printed with --details. The label is
    the text in column B of that row."""

    def __init__(self, sheet, title, compare="formulas"):
        self.sheet, self.title = sheet, title
        self.cells = {}           # "D37" -> value
        self.labels = {}          # row -> text expected in column B (checked before comparing that row)
        self.headers = {}         # "W5" -> text expected in that header cell (workings sheets)
        self.panels = []          # (heading, [(column letter, header)], [lines])
        self.compare = compare    # the workbook check compared every formula cell ("formulas") or only these

    def panel(self, heading, columns):
        self.panels.append((heading, list(columns.items()), []))

    def row(self, r, label, fmt, check_label=True, **values):
        for c, v in values.items():
            self.cells[f"{c}{r}"] = v
        if check_label and label is not None:
            self.labels[r] = label
        self.panels[-1][2].append(("row", label, fmt, values))

    def text(self, cell, value):
        """A text cell kept with the values but not printed as a table row (column headers, short descriptions)."""
        self.cells[cell] = value

    def note(self, line):
        self.panels[-1][2].append(("note", line, None, None))


def fmt_value(v, fmt):
    if isinstance(v, str):
        return v
    if isinstance(v, (dt.date, pd.Timestamp, np.datetime64)):
        return to_date(v).isoformat()
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    if fmt == "pct":
        return f"{v * 100:.2f}%"
    if fmt == "pct3":
        return f"{v * 100:.3f}%"
    if fmt == "year":
        return f"{v:.0f}"
    if fmt == "int":
        return f"{v:,.0f}"
    if fmt == "p":
        return f"{v:.4f}" if v >= 0.0001 or v == 0 else f"{v:.2e}"
    if fmt == "usd":
        return f"{v:,.2f}"
    if fmt == "e":
        return f"{v:.2e}"
    if fmt == "num2":
        return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}"
    return f"{v:.4f}"


def pick_fmt(fmt, c):
    """A row's number format: one code for every column, or a dict {column: code}."""
    return fmt.get(c, "") if isinstance(fmt, dict) else fmt


def print_report(rep, width=140):
    print("\n" + "=" * width)
    print(f"{rep.sheet.upper()}  |  {rep.title}")
    print("=" * width)
    label_w = 66
    for heading, columns, lines in rep.panels:
        print(f"\n--- {heading} ---")
        rows = [ln for ln in lines if ln[0] == "row"]
        colw = {}
        for c, h in columns:
            vals = [fmt_value(ln[3][c], pick_fmt(ln[2], c)) for ln in rows if c in ln[3]]
            vals = [v for v in vals if len(v) <= 40]
            colw[c] = max([len(h)] + [len(v) for v in vals] + [10]) + 2
        if columns:
            print(" " * label_w + "".join(h.rjust(colw[c]) for c, h in columns))
        for kind, label, fmt, values in lines:
            if kind == "note":
                print("  " + label)
                continue
            lab = str(label if label is not None else "")
            lab = lab if len(lab) <= label_w - 2 else lab[:label_w - 3] + "…"
            cells = []
            long_texts = []
            for c, h in columns:
                if c in values:
                    s = fmt_value(values[c], pick_fmt(fmt, c))
                    if len(s) > 40:
                        long_texts.append(s)
                        s = "(see below)"
                    cells.append(s.rjust(colw[c]))
                else:
                    cells.append(" " * colw[c])
            print(lab.ljust(label_w) + "".join(cells))
            for s in long_texts:
                print("    " + s)

# ---------------------------------------------------------------------------------------------------------------------
# 2.7 Settings sheet
# ---------------------------------------------------------------------------------------------------------------------
def report_settings(da):
    rep = Report("Settings", "every assumption used by the analysis (named constants at the top of this script)",
                 compare="listed")
    rep.panel("Inputs and derived constants", {"C": "Value"})
    rows = [
        (7, "Significance level (two-sided)", "ALPHA", ALPHA, "num"),
        (8, "Critical value for that level (normal)", "ZCRIT", ZCRIT, "num"),
        (9, "Power for the minimum detectable effect", "POWER", POWER, "num"),
        (10, "z for that power (normal)", "ZPOW", ZPOW, "num"),
        (11, "MDE multiplier", "MDE_MULT", MDE_MULT, "num"),
        (12, "Smallest effect worth detecting, per trade", "MIN_EFFECT", MIN_EFFECT, "num"),
        (13, "Trading days per year", "DAYS_PER_YEAR", DAYS_PER_YEAR, "int"),
        (14, "Minimum days needed to estimate a beta", "MIN_BETA_DAYS", MIN_BETA_DAYS, "int"),
        (15, "Newey–West lags", "NW_LAGS", NW_LAGS, "int"),
        (16, "Last year of the first half", "SPLIT_YEAR", SPLIT_YEAR, "year"),
        (17, "Winsorising: lower percentile (V3)", "WINSOR_LO", WINSOR_LO, "num"),
        (18, "Winsorising: upper percentile (V3)", "WINSOR_HI", WINSOR_HI, "num"),
        (19, "Trades dropped at each end (V6)", "TRIM_N", TRIM_N, "int"),
        (20, "Year left out in V4 (first)", "EXCL_YEAR_1", EXCL_YEAR_1, "year"),
        (21, "Year left out in V4 (second)", "EXCL_YEAR_2", EXCL_YEAR_2, "year"),
        (22, "Last day with Fama-French factors", "FF_LAST (from the data)", to_date(da.attrs["ff_last"]), "num"),
        (23, "Start of the 'from 2020' check", "START_2020", START_2020, "num"),
        (24, "Break test: share of dates kept out at each end", "BREAK_TRIM", BREAK_TRIM, "num"),
        (25, "Break test: critical value at 5%", "BAI_PERRON_CV_5", BAI_PERRON_CV_5, "num"),
        (26, "Break test: critical value at 1%", "BAI_PERRON_CV_1", BAI_PERRON_CV_1, "num"),
        (27, "Largest quarterly trades left out (E5)", "N_DROP_V1", N_DROP_V1, "int"),
    ]
    for r, label, name, value, fmt in rows:
        rep.row(r, f"{label}  [{name}]", fmt, check_label=False, C=value)
        rep.labels[r] = label
    rep.note(f"Worked example in Tests 1 and 2: {WORKED_EXAMPLE}. Beta window: days -{BETA_FIRST_DAY} to "
             f"-{BETA_LAST_DAY}. The account starts on the first trading day on or after {text_dmy(ACCOUNT_START)}.")
    return rep


# ---------------------------------------------------------------------------------------------------------------------
# 2.8 Workings summaries (Trade Returns, Beta Estimation, Holding Days, Daily Account)
# ---------------------------------------------------------------------------------------------------------------------
def print_workings(raw, tr, beta, hd, da):
    print("\n" + "=" * 140)
    print("WORKINGS  |  Trade Returns, Beta Estimation, Holding Days, Daily Account "
          "(--save writes them out in full)")
    print("=" * 140)
    print(f"Trades: {raw.n} additions, announced "
          f"{tr.announced.min():%d %b %Y} to {tr.announced.max():%d %b %Y}, on {tr.announced.nunique()} dates; "
          f"{(tr.quarterly == 'Yes').sum()} at a quarterly rebalance.")
    print(f"Holding days: {len(hd):,} position-days (median {tr.days.median():.0f} days, range "
          f"{tr.days.min()}-{tr.days.max()}).")
    few = beta.days.values < MIN_BETA_DAYS
    print(f"Beta: {(~few).sum()} estimated, {few.sum()} set to 1 "
          f"({', '.join(f'{n} ({d} days)' for n, d in zip(tr.name.values[few], beta.days.values[few]))}).")
    print(f"Daily Account: {len(da):,} trading days, {to_date(da.date.iat[0])} to {to_date(da.date.iat[-1])}; "
          f"Fama-French factors to {to_date(da.attrs['ff_last'])}.")
    show = tr[["row", "name", "announced", "included", "days", "AR_A", "AR_B", "beta", "AR_A_adj", "AR_B_adj", "V5"]]
    print("\nTrade Returns, the 5 largest and 5 smallest Window A abnormal returns:")
    pick = pd.concat([show.nlargest(5, "AR_A"), show.nsmallest(5, "AR_A")])
    with pd.option_context("display.width", 160, "display.max_columns", 20, "display.float_format", "{:.4f}".format):
        print(pick.to_string(index=False))


def save_workings(folder, tr, beta, hd, da):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    tr.assign(beta_days=beta.days.values, beta_slope=beta.slope.values).to_csv(folder / "trade_returns.csv",
                                                                               index=False)
    hd.to_csv(folder / "holding_days.csv", index=False)
    da.to_csv(folder / "daily_account.csv", index=False)
    print(f"\nWorkings saved in {folder}: trade_returns.csv, holding_days.csv, daily_account.csv")


# ---------------------------------------------------------------------------------------------------------------------
# 2.9 Test 1 Per Trade and Test 2 Risk-Adjusted
# ---------------------------------------------------------------------------------------------------------------------
def per_trade_sections(rep, a, b, test, off, beat_word):
    """Sections a-f and the checks of Test 1 (off = 0) or Test 2 (off = 12; checks at 116 instead of 135).
    a, b = stat blocks of Window A (column D) and Window B (column E)."""
    sig5 = f"Significant at the {text_pct(ALPHA)} level?"
    rep.panel(f"{test}a. Average abnormal return, each trade once", {"D": "Window A", "E": "Window B"})
    rep.row(36 + off, "N: number of trades", "int", D=a.N, E=b.N)
    rep.row(37 + off, "Mean abnormal return", "pct", D=a.mean, E=b.mean)
    rep.row(38 + off, "s: standard deviation", "pct", D=a.s, E=b.s)
    rep.row(39 + off, "SE: standard error of the mean", "pct", D=a.se, E=b.se)
    rep.row(40 + off, "t-statistic", "num", D=a.t, E=b.t)
    rep.row(41 + off, "p-value (two-sided)", "p", D=a.p, E=b.p)
    rep.row(42 + off, sig5, "", D=sig_word(a.p, a.mean), E=sig_word(b.p, b.mean))
    rep.panel(f"{test}b. Clustered by announcement date (each date counts once)", {"D": "Window A", "E": "Window B"})
    rep.row(50 + off, "G: number of announcement dates", "int", D=a.G, E=b.G)
    rep.row(51 + off, "…of which dates with 2 or more trades", "int", D=a.multi, E=b.multi)
    rep.row(52 + off, "Mean of the date averages", "pct", D=a.meanG, E=b.meanG)
    rep.row(53 + off, "s of the date averages", "pct", D=a.sdG, E=b.sdG)
    rep.row(54 + off, "SE (clustered)", "pct", D=a.seG, E=b.seG)
    rep.row(55 + off, "t-statistic (clustered)", "num", D=a.tG, E=b.tG)
    rep.row(56 + off, "p-value (two-sided)", "p", D=a.pG, E=b.pG)
    rep.row(57 + off, sig5, "", D=sig_word(a.pG, a.meanG), E=sig_word(b.pG, b.meanG))
    rep.panel(f"{test}c. {text_pct(1 - ALPHA)} confidence interval for the mean of the date averages",
              {"D": "Window A", "E": "Window B"})
    rep.row(62 + off, f"t*: the t value that leaves {text_pct(ALPHA / 2, 1)} in each tail", "num",
            D=a.t_star, E=b.t_star)
    rep.text(f"C{62 + off}", f"T.INV.2T({text_pct(ALPHA)}, G − 1)")
    rep.row(63 + off, "Lower end", "pct", D=a.lo, E=b.lo)
    rep.row(64 + off, "Upper end", "pct", D=a.hi, E=b.hi)
    rep.row(65 + off, "Does the interval exclude zero?", "", D=yes(a.lo > 0 or a.hi < 0), E=yes(b.lo > 0 or b.hi < 0))
    rep.panel(f"{test}d. Median and Wilcoxon signed-rank test", {"D": "Window A", "E": "Window B"})
    rep.row(72 + off, "Median abnormal return", "pct", D=a.median, E=b.median)
    rep.row(73 + off, "n: trades with a non-zero return", "int", D=a.n, E=b.n)
    rep.row(74 + off, "W+: sum of the ranks of the positive trades", "num2", D=a.W, E=b.W)
    rep.row(75 + off, "Expected W+ if the strategy does nothing", "num2", D=a.EW, E=b.EW)
    rep.row(76 + off, "Variance of W+ if the strategy does nothing", "num2", D=a.VW, E=b.VW)
    rep.row(77 + off, "z-statistic", "num", D=a.z, E=b.z)
    rep.row(78 + off, "p-value (two-sided)", "p", D=a.pW, E=b.pW)
    rep.row(79 + off, sig5, "", D=sig_word(a.pW, a.z), E=sig_word(b.pW, b.z))
    rep.panel(f"{test}e. Sign test", {"D": "Window A", "E": "Window B"})
    beat50 = lambda s: ("Yes, above 50%" if s.beat > 0.5 else "Yes, below 50%") if s.pS < ALPHA else "No"
    rep.row(84 + off, f"k: trades that beat {beat_word}", "int", D=a.k, E=b.k)
    rep.row(85 + off, "n: trades with a non-zero return", "int", D=a.n, E=b.n)
    rep.row(86 + off, f"% of trades that beat {beat_word}", "pct", D=a.beat, E=b.beat)
    rep.row(87 + off, "p-value (two-sided sign test)", "p", D=a.pS, E=b.pS)
    rep.row(88 + off, "Significantly different from 50%?", "", D=beat50(a), E=beat50(b))
    rep.panel(f"{test}f. Minimum detectable effect (power {text_pct(POWER)})", {"D": "Window A", "E": "Window B"})
    rep.row(93 + off, f"SE (clustered, from {test}b)", "pct", D=a.seG, E=b.seG)
    rep.row(94 + off, "MDE: minimum detectable effect", "pct", D=a.mde, E=b.mde)
    rep.row(95 + off, "Smallest effect worth detecting", "pct", D=MIN_EFFECT, E=MIN_EFFECT)
    rep.row(96 + off, "Could the test detect it?", "", D=yes(a.mde <= MIN_EFFECT), E=yes(b.mde <= MIN_EFFECT))


def verdict_checks(rep, a, b, r0, test):
    """The six checks behind Test 1's rule, and the verdict, rows r0..r0+6."""
    rep.panel("Putting it together (Test 1's rule)", {"D": "Window A", "E": "Window B"})
    rep.row(r0, "1. Average above zero?", "", D=yes(a.meanG > 0), E=yes(b.meanG > 0))
    rep.row(r0 + 1, "2. …and significantly so?", "", D=yes(a.pG < ALPHA and a.meanG > 0),
            E=yes(b.pG < ALPHA and b.meanG > 0))
    rep.text(f"C{r0 + 1}", f"{test}b: clustered p below {text_pct(ALPHA)} with the mean above 0")
    rep.row(r0 + 2, f"3. {text_pct(1 - ALPHA)} interval entirely above zero?", "", D=yes(a.lo > 0), E=yes(b.lo > 0))
    rep.row(r0 + 3, "4. Median above zero?", "", D=yes(a.median > 0), E=yes(b.median > 0))
    rep.row(r0 + 4, "5. …and significantly so?", "", D=yes(a.pW < ALPHA and a.z > 0), E=yes(b.pW < ALPHA and b.z > 0))
    rep.text(f"C{r0 + 4}", f"{test}d: Wilcoxon p below {text_pct(ALPHA)} with z above 0")
    rep.row(r0 + 5, f"6. Could the test detect {text_pct(MIN_EFFECT)} per trade?", "",
            D=yes(a.mde <= MIN_EFFECT), E=yes(b.mde <= MIN_EFFECT))
    rep.row(r0 + 6, "Result for this window", "", D=verdict(a), E=verdict(b))


def worked_example(rep, tr, r0, risk_adjusted=False):
    """The worked example (one trade, both windows), rows r0..r0+9 (and +10, +11 for the beta lines in Test 2)."""
    i = int(np.nonzero(tr.label.values == WORKED_EXAMPLE)[0][0])
    t = tr.iloc[i]
    rep.panel(f"Worked example: {WORKED_EXAMPLE}", {"D": "Window A", "E": "Window B"})
    rep.row(r0, "Number in the list of additions (1 = newest)", "int", D=int(t.row))
    rep.row(r0 + 1, "Announced", "", D=to_date(t.announced))
    rep.row(r0 + 2, "Included (first day in the index)", "", D=to_date(t.included))
    rep.row(r0 + 3, "Stock: buy at the Ann+1 open ($)", "usd", D=t.s_buy_open, E=t.s_buy_open)
    rep.row(r0 + 4, "Stock: sell ($)", "usd", D=t.s_inc_open, E=t.s_inc_close)
    rep.row(r0 + 5, "Stock return R", "pct", D=t.R_A, E=t.R_B)
    rep.row(r0 + 6, "SPY: buy at the Ann+1 open ($)", "usd", D=t.m_buy_open, E=t.m_buy_open)
    rep.row(r0 + 7, "SPY: sell ($)", "usd", D=t.m_inc_open, E=t.m_inc_close)
    rep.row(r0 + 8, "SPY return M", "pct", D=t.M_A, E=t.M_B)
    rep.row(r0 + 9, "Abnormal return AR = R − M", "pct", D=t.AR_A, E=t.AR_B)
    if risk_adjusted:
        rep.row(r0 + 10, "Beta (Beta Estimation sheet)", "num", D=t.beta)
        rep.row(r0 + 11, "Risk-adjusted AR = R − β × M", "pct", D=t.AR_A_adj, E=t.AR_B_adj)


def report_test1(tr):
    rep = Report("Test 1 Per Trade",
                 "does buying at the Ann+1 open and selling at inclusion beat SPY over the same hours?")
    dates = tr.announced.values
    a, b = stat_block(tr.AR_A, dates), stat_block(tr.AR_B, dates)
    rep.panel("Answer", {"D": "Result"})
    rep.row(9, "Test 1 result", "", D=combine(verdict(a), verdict(b)))
    worked_example(rep, tr, 22)
    per_trade_sections(rep, a, b, "1", 0, "SPY")
    # Figure 1: average abnormal return since the announcement close, at six moments
    rep.panel("Figure 1. Average abnormal return since the announcement close (each date once; band = mean ± 1.96 SE)",
              {"D": "Mean (dates)", "E": "SE", "F": f"Lower {text_pct(1 - ALPHA)}", "G": f"Upper {text_pct(1 - ALPHA)}",
               "H": "Mean (trades)"})
    rep.text("F102", f"Lower {text_pct(1 - ALPHA)}")
    rep.text("G102", f"Upper {text_pct(1 - ALPHA)}")
    moments = [(103, "Announcement close", "CAR_ann_close"), (104, "Ann+1 open (we buy)", "CAR_buy_open"),
               (105, "Ann+1 close", "CAR_ann1_close"), (106, "Close the day before inclusion", "CAR_pre_close"),
               (107, "Inclusion open (sell, Window A)", "CAR_inc_open"),
               (108, "Inclusion close (sell, Window B)", "CAR_inc_close")]
    for r, label, c in moments:
        by_date = tr.groupby("announced")[c].mean()
        m, se = by_date.mean(), stdev_s(by_date.values) / math.sqrt(a.G)
        rep.row(r, label, "pct", D=m, E=se, F=m - ZCRIT * se, G=m + ZCRIT * se, H=tr[c].mean())
    verdict_checks(rep, a, b, 135, "1")
    return rep, a, b


def report_test2(tr, beta, t1a, t1b):
    rep = Report("Test 2 Risk-Adjusted", "the same tests on R − β × M (beta from days −250 to −30)")
    dates = tr.announced.values
    a, b = stat_block(tr.AR_A_adj, dates), stat_block(tr.AR_B_adj, dates)
    va, vb = verdict(a), verdict(b)
    same_a, same_b = yes(verdict(t1a) == va), yes(verdict(t1b) == vb)
    rep.panel("Answer", {"D": "Result"})
    rep.row(9, "Test 2 result", "", D=combine(va, vb))
    rep.row(10, "Same answer as Test 1?", "", D=yes(same_a == "Yes" and same_b == "Yes"))
    rep.panel("Betas", {"D": "Value"})
    used = beta.beta.values
    rep.row(20, "Trades with an estimated beta", "int", D=int((beta.days >= MIN_BETA_DAYS).sum()))
    rep.row(21, "Trades with beta set to 1", "int", D=int((beta.days < MIN_BETA_DAYS).sum()))
    rep.row(22, "Average beta used", "num", D=used.mean())
    rep.row(23, "Median beta used", "num", D=float(np.median(used)))
    rep.row(24, "Lowest beta used", "num", D=used.min())
    rep.row(25, "Highest beta used", "num", D=used.max())
    rep.row(26, "Share of stocks with a beta above 1", "pct", D=(used > 1).sum() / len(used))
    worked_example(rep, tr, 32, risk_adjusted=True)
    per_trade_sections(rep, a, b, "2", 12, "β × SPY")
    verdict_checks(rep, a, b, 116, "2")
    rep.panel("Compared with Test 1", {"D": "Window A", "E": "Window B"})
    rep.row(126, "Average, each date once: Test 1", "pct", D=t1a.meanG, E=t1b.meanG)
    rep.row(127, "Average, each date once: risk-adjusted", "pct", D=a.meanG, E=b.meanG)
    rep.row(128, "Difference", "pct", D=a.meanG - t1a.meanG, E=b.meanG - t1b.meanG)
    rep.row(129, "Average, each trade once: Test 1", "pct", D=t1a.mean, E=t1b.mean)
    rep.row(130, "Average, each trade once: risk-adjusted", "pct", D=a.mean, E=b.mean)
    rep.row(131, "Result, Test 1", "", D=verdict(t1a), E=verdict(t1b))
    rep.row(132, "Result, risk-adjusted", "", D=va, E=vb)
    rep.row(133, "Same answer as Test 1?", "", D=same_a, E=same_b)
    return rep


# ---------------------------------------------------------------------------------------------------------------------
# 2.10 Test 3 Portfolio
# ---------------------------------------------------------------------------------------------------------------------
def account_panel(rep, da, k, off, regs, z_last):
    """Panel A (k = 'A', off = 0) or Panel B (k = 'B', off = 42): the Long + SPY account (column D) against B1 (E)."""
    T = len(da)
    res = {}
    for col, key in [("D", k), ("E", "B1")]:
        r, x, v = da[key].values, da["x_" + key].values, da["value_" + key].values
        res[col] = SimpleNamespace(mean=r.mean(), ann=v[-1] ** (DAYS_PER_YEAR / T) - 1, end=v[-1],
                                   vol=stdev_s(r) * math.sqrt(DAYS_PER_YEAR), xmean=x.mean(), xsd=stdev_s(x),
                                   srd=x.mean() / stdev_s(x), dd=drawdown(da, key))
    D, E = res["D"], res["E"]
    jk = jobson_korkie(da["x_" + k].values, da.x_B1.values)
    capm, ff4 = regs
    name = "Long + SPY, Window " + k
    rep.panel(f"Panel {k}: {name} (D) against B1, SPY bought and held (E)", {"D": name, "E": "B1"})
    rep.row(34 + off, "Days (T)", "int", D=T, E=T)
    rep.row(35 + off, "% of days holding added stocks", "pct", D=(da.positions > 0).sum() / T)
    rep.row(36 + off, "Average daily return", "pct3", D=D.mean, E=E.mean)
    rep.row(37 + off, "Annualised return", "pct", D=D.ann, E=E.ann)
    rep.text(f"C{37 + off}", f"(value of $1 at the end)^({DAYS_PER_YEAR} ÷ T) − 1")
    rep.row(38 + off, "Return per year above B1", "pct", D=D.ann - E.ann)
    rep.row(39 + off, "Value of $1 at the end", "usd", D=D.end, E=E.end)
    rep.row(40 + off, "Annualised volatility", "pct", D=D.vol, E=E.vol)
    rep.text(f"C{40 + off}", f"STDEV.S of the daily returns × √{DAYS_PER_YEAR}")
    rep.row(41 + off, "Average daily excess return", "pct3", D=D.xmean, E=E.xmean)
    rep.row(42 + off, "St. dev. of daily excess return", "pct3", D=D.xsd, E=E.xsd)
    rep.row(43 + off, "Sharpe ratio, daily", "num", D=D.srd, E=E.srd)
    rep.row(44 + off, "Sharpe ratio, annualised", "num", D=D.srd * math.sqrt(DAYS_PER_YEAR),
            E=E.srd * math.sqrt(DAYS_PER_YEAR))
    rep.text(f"C{44 + off}", f"Daily Sharpe × √{DAYS_PER_YEAR}")
    rep.row(46 + off, "ρ: correlation with B1", "num", D=jk.rho)
    rep.row(47 + off, "SE of the Sharpe difference", "num", D=jk.se)
    rep.row(48 + off, "z-statistic", "num", D=jk.z)
    rep.row(49 + off, "p-value (two-sided)", "p", D=jk.p)
    above, below = yes(jk.z > 0 and jk.p < ALPHA), yes(jk.z < 0 and jk.p < ALPHA)
    rep.row(50 + off, "Significantly above B1?", "", D=above)
    rep.row(51 + off, "Significantly below B1?", "", D=below)
    rep.row(54 + off, "Days in the regressions", "int", D=z_last)
    rep.row(55 + off, "CAPM alpha, per year", "pct", D=capm.alpha_year)
    rep.text(f"C{55 + off}", f"Daily α × {DAYS_PER_YEAR}")
    rep.row(56 + off, "t-statistic (Newey–West)", "num", D=capm.t)
    rep.row(57 + off, "p-value (two-sided)", "p", D=t_dist_2t(capm.t, z_last - 2))
    rep.row(58 + off, "CAPM beta", "num", D=capm.b[1])
    rep.row(59 + off, "Four-factor alpha, per year", "pct", D=ff4.alpha_year)
    rep.text(f"C{59 + off}", f"Daily α × {DAYS_PER_YEAR}, with Mkt-RF, SMB, HML and Mom")
    rep.row(60 + off, "t-statistic (Newey–West)", "num", D=ff4.t)
    rep.row(61 + off, "p-value (two-sided)", "p", D=t_dist_2t(ff4.t, z_last - 5))
    rep.row(62 + off, "Loading on Mkt-RF", "num", D=ff4.b[1])
    rep.row(63 + off, "Loading on SMB (small stocks)", "num", D=ff4.b[2])
    rep.row(64 + off, "Loading on HML (value stocks)", "num", D=ff4.b[3])
    rep.row(65 + off, "Loading on Mom (recent winners)", "num", D=ff4.b[4])
    rep.row(68 + off, "Maximum drawdown", "pct", D=D.dd.max_dd, E=E.dd.max_dd)
    rep.row(69 + off, "Peak before the fall", "", D=D.dd.peak, E=E.dd.peak)
    rep.row(70 + off, "Trough", "", D=D.dd.trough, E=E.dd.trough)
    rep.row(71 + off, "Recovered by", "", D=D.dd.recovered, E=E.dd.recovered)
    return SimpleNamespace(sharpe_above=above, capm=yes(capm.alpha_year > 0 and capm.t > ZCRIT),
                           ff4=yes(ff4.alpha_year > 0 and ff4.t > ZCRIT))


def report_test3(tr, hd, da, regs):
    rep = Report("Test 3 Portfolio", "the strategy run as an account (Long + SPY) against SPY bought and held (B1)")
    T = len(da)
    z_last = regs["full_T"]
    no_pos = da.positions.values == 0
    rep.panel("Checks on the workings", {"D": "Value", "E": "OK?"})
    rep.row(24, "Position-days in the Holding Days sheet", "int", D=len(hd), E=yes(len(hd) == int(tr.days.sum())))
    rep.row(25, f"Trading days held, summed over the {len(tr)} trades", "int", D=int(tr.days.sum()))
    rep.row(26, "Positions held, summed over every day of the account", "int", D=int(da.positions.sum()),
            E=yes(int(da.positions.sum()) == len(hd)))
    g_stock = np.abs(tr[["gap_stock_A", "gap_stock_B"]].values).max()
    g_spy = np.abs(tr[["gap_spy_A", "gap_spy_B"]].values).max()
    rep.row(27, "Largest gap between a stock's chained pieces and its trade return", "e", D=g_stock,
            E=yes(g_stock < 1e-6))
    rep.row(28, "Largest gap between SPY's chained pieces and its window return", "e", D=g_spy, E=yes(g_spy < 1e-6))
    bad = int((no_pos & (np.abs(da.A.values - da.B1.values) > 1e-7)).sum()
              + (no_pos & (np.abs(da.B.values - da.B1.values) > 1e-7)).sum())
    rep.row(29, "Long + SPY equals B1 on every day with no position", "int", D=bad, E=yes(bad == 0))
    found = int(da.price_found.sum())
    rep.row(30, "Days whose SPY price matches the date", "", D=f"{found:,} of {T:,}", E=yes(found == T))
    A = account_panel(rep, da, "A", 0, (regs["capm_A"], regs["ff4_A"]), z_last)
    B = account_panel(rep, da, "B", 42, (regs["capm_B"], regs["ff4_B"]), z_last)
    rep.panel("Putting it together", {"D": "Window A", "E": "Window B"})
    rep.row(120, "Sharpe ratio reliably above B1?", "", D=A.sharpe_above, E=B.sharpe_above)
    rep.row(121, f"CAPM alpha above zero with t > {ZCRIT:.2f}?", "", D=A.capm, E=B.capm)
    rep.row(122, f"Four-factor alpha above zero with t > {ZCRIT:.2f}?", "", D=A.ff4, E=B.ff4)
    res = {c: ("Works" if (w.sharpe_above, w.capm, w.ff4) == ("Yes",) * 3 else "Doesn't work") for c, w in
           [("D", A), ("E", B)]}
    rep.row(123, "Result for this window", "", D=res["D"], E=res["E"])
    rep.panel("Answer", {"D": "Result"})
    rep.row(8, "Test 3 result", "", D="Works" if "Works" in res.values() else "Doesn't work")
    rep.panels.insert(0, rep.panels.pop())
    return rep, res


# ---------------------------------------------------------------------------------------------------------------------
# 2.11 Test 4 Robustness
# ---------------------------------------------------------------------------------------------------------------------
STAT_ROWS = ["N: trades", "Mean abnormal return (each trade once)", "Median", "% of trades that beat SPY",
             "Announcement dates (G)", "Mean of the date averages (each date once)", "Clustered SE",
             "Clustered t = mean of date averages ÷ SE", "p-value: T.DIST.2T(|t|, G − 1)",
             None, "Wilcoxon n (non-zero trades)", "W+: sum of ranks of positive trades", "Expected W+ = n(n + 1) ÷ 4",
             "z = (W+ − expected) ÷ √[n(n + 1)(2n + 1) ÷ 24]", "Wilcoxon p-value", None,
             "Average significant? (clustered t-test)", "Verdict (Test 1's rule)", "Same verdict as Test 1?"]


def robust_block(rep, r0, heading, columns, test1_verdicts):
    """Test 4's statistic block (rows r0..r0+18). columns = {letter: (header, x, mask, dates)}; the statistics use
    only the trades in mask, with date averages and Wilcoxon ranks computed within that subset."""
    blocks = {c: stat_block(x[m], dates[m]) for c, (_, x, m, dates) in columns.items()}
    rep.panel(heading, {c: v[0] for c, v in columns.items()})
    labels = list(STAT_ROWS)
    labels[9] = f"Lower end of the {text_pct(1 - ALPHA)} interval: mean − t* × SE"
    labels[15] = f"MDE = ({ZCRIT:.2f} + {ZPOW:.2f}) × clustered SE"
    fields = [("N", "int"), ("mean", "pct"), ("median", "pct"), ("beat", "pct"), ("G", "int"), ("meanG", "pct"),
              ("seG", "pct"), ("tG", "num"), ("pG", "p"), ("lo", "pct"), ("n", "int"), ("W", "num2"),
              ("EW", "num2"), ("z", "num"), ("pW", "p"), ("mde", "pct")]
    for j, (f, fmt) in enumerate(fields):
        rep.row(r0 + j, labels[j], fmt, **{c: getattr(b, f) for c, b in blocks.items()})
    rep.row(r0 + 16, labels[16], "", **{c: sig_word(b.pG, b.meanG) for c, b in blocks.items()})
    rep.row(r0 + 17, labels[17], "", **{c: verdict(b) for c, b in blocks.items()})
    rep.row(r0 + 18, labels[18], "", **{c: yes(verdict(b) == test1_verdicts[c]) for c, b in blocks.items()})
    return blocks


def report_test4(tr, da, regs, t1a, t1b):
    rep = Report("Test 4 Robustness", "does the Test 1 answer hold in each half and in variants of the sample?")
    d = tr.announced.values
    xA, xB, xV5 = tr.AR_A.values, tr.AR_B.values, tr.V5.values
    year = tr.year.values
    vA, vB = verdict(t1a), verdict(t1b)
    y0, y1 = int(year.min()), int(year.max())
    H1, H2 = year <= SPLIT_YEAR, year > SPLIT_YEAR
    ALL = np.ones(len(tr), bool)
    Q, OFF = tr.quarterly.values == "Yes", tr.quarterly.values == "No"
    V4 = ~np.isin(year, [EXCL_YEAR_1, EXCL_YEAR_2])
    # V3: winsorise at the 1st and 99th percentiles (PERCENTILE.INC = numpy's default linear interpolation)
    loA, hiA = np.percentile(xA, 100 * WINSOR_LO), np.percentile(xA, 100 * WINSOR_HI)
    loB, hiB = np.percentile(xB, 100 * WINSOR_LO), np.percentile(xB, 100 * WINSOR_HI)
    wA, wB = np.minimum(np.maximum(xA, loA), hiA), np.minimum(np.maximum(xB, loB), hiB)
    # V6: drop the TRIM_N largest and TRIM_N smallest (RANK.AVG from the top and from the bottom both > TRIM_N)
    trim = lambda x: (rank_avg(-x) > TRIM_N) & (rank_avg(x) > TRIM_N)
    V6A, V6B = trim(xA), trim(xB)

    # 4a: first and second half
    h = {"C": f"First half, A ({y0}–{SPLIT_YEAR})", "D": f"Second half, A ({SPLIT_YEAR + 1}–{y1})",
         "E": f"First half, B ({y0}–{SPLIT_YEAR})", "F": f"Second half, B ({SPLIT_YEAR + 1}–{y1})"}
    for c, lab in h.items():
        rep.text(f"{c}17", lab)
    halves = robust_block(rep, 18, "4a. First half and second half (by announcement year)",
                          {"C": (h["C"], xA, H1, d), "D": (h["D"], xA, H2, d), "E": (h["E"], xB, H1, d),
                           "F": (h["F"], xB, H2, d)}, {"C": vA, "D": vA, "E": vB, "F": vB})
    # Do the halves differ? Welch t-test and Mann-Whitney
    rep.panel("4a. Do the halves differ? (Welch t-test and Mann–Whitney)", {"C": "Window A", "D": "Window B"})
    res = {}
    for c, x in [("C", xA), ("D", xB)]:
        a, b = x[H1], x[H2]
        n1, n2 = len(a), len(b)
        s1, s2 = stdev_s(a), stdev_s(b)
        se = math.sqrt(s1 ** 2 / n1 + s2 ** 2 / n2)
        diff = b.mean() - a.mean()
        df = (s1 ** 2 / n1 + s2 ** 2 / n2) ** 2 / ((s1 ** 2 / n1) ** 2 / (n1 - 1) + (s2 ** 2 / n2) ** 2 / (n2 - 1))
        R1 = rank_avg(x)[H1].sum()                       # ranks among all trades of this window
        U = R1 - n1 * (n1 + 1) / 2
        EU = n1 * n2 / 2
        zU = (U - EU) / math.sqrt(n1 * n2 * (n1 + n2 + 1) / 12)
        res[c] = dict(n1=n1, n2=n2, m1=a.mean(), m2=b.mean(), s1=s1, s2=s2, diff=diff, se=se, t=diff / se, df=df,
                      p=t_dist_2t(diff / se, df), R1=R1, U=U, EU=EU, z=zU, pU=norm_2t(zU))
    rows = [(40, "n, first half", "n1", "int"), (41, "n, second half", "n2", "int"),
            (42, "Mean, first half", "m1", "pct"), (43, "Mean, second half", "m2", "pct"),
            (44, "St. dev., first half", "s1", "pct"), (45, "St. dev., second half", "s2", "pct"),
            (46, "Difference (second − first)", "diff", "pct"),
            (47, "SE of the difference = √(s₁²/n₁ + s₂²/n₂)", "se", "pct"),
            (48, "Welch t = difference ÷ SE", "t", "num"), (49, "Welch degrees of freedom", "df", "num2"),
            (50, "Welch p-value: T.DIST.2T(|t|, df)", "p", "p"),
            (51, "Mann–Whitney: sum of ranks of the first half", "R1", "num2"),
            (52, "U = sum of ranks − n₁(n₁ + 1) ÷ 2", "U", "num2"),
            (53, "Expected U if the halves are alike = n₁n₂ ÷ 2", "EU", "num2"),
            (54, "z = (U − expected) ÷ √[n₁n₂(n₁ + n₂ + 1) ÷ 12]", "z", "num"),
            (55, "Mann–Whitney p-value", "pU", "p")]
    for r, label, key, fmt in rows:
        rep.row(r, label, fmt, C=res["C"][key], D=res["D"][key])
    rep.row(56, "Do the halves differ significantly?", "",
            **{c: yes(res[c]["p"] < ALPHA or res[c]["pU"] < ALPHA) for c in "CD"})

    # The account by half: Sharpe ratios (to the end of the account) and CAPM alphas (to French's last day)
    hday, T = regs["split_day"], len(da)
    lab_c = f"First half (to {text_dmy(da.date.iat[hday - 1])})"
    lab_d = f"Second half (from {text_dmy(da.date.iat[hday])})"
    rep.text("C60", lab_c)
    rep.text("D60", lab_d)
    rep.panel("4a. The account by half", {"C": lab_c, "D": lab_d})
    sh = lambda k, a, b: sharpe_daily(da["x_" + k].values[a:b]) * math.sqrt(DAYS_PER_YEAR)
    rep.row(61, "Sharpe ratio, Long + SPY, Window A", "num", C=sh("A", 0, hday), D=sh("A", hday, T))
    rep.row(62, "Sharpe ratio, Long + SPY, Window B", "num", C=sh("B", 0, hday), D=sh("B", hday, T))
    rep.row(63, "Sharpe ratio, B1 SPY buy and hold", "num", C=sh("B1", 0, hday), D=sh("B1", hday, T))
    for r, k in [(64, "A"), (68, "B")]:
        r1, r2 = regs[f"capm_{k}_h1"], regs[f"capm_{k}_h2"]
        p1, p2 = t_dist_2t(r1.t, r1.T - 2), t_dist_2t(r2.t, r2.T - 2)
        rep.row(r, f"CAPM alpha per year, Long + SPY account, Window {k}", "pct", C=r1.alpha_year, D=r2.alpha_year)
        rep.row(r + 1, "…t-statistic (Newey–West)", "num", C=r1.t, D=r2.t)
        rep.row(r + 2, "…p-value: T.DIST.2T(|t|, days − 2)", "p", C=p1, D=p2)
        rep.row(r + 3, f"Significant at the {text_pct(ALPHA)} level?", "", C=sig_word(p1, r1.alpha_year),
                D=sig_word(p2, r2.alpha_year))

    # 4b: year by year
    rep.panel("4b. Year by year", {"B": "Year", "C": "Trades", "D": "Mean A", "E": "Median A", "F": "Beat A",
                                   "G": "Mean B", "H": "Median B", "I": "Beat B", "J": "Label"})
    beat = lambda x: (x > 0).sum() / (x != 0).sum()
    for y in range(y0, y1 + 1):
        m = year == y
        r = 77 + (y - y0)
        rep.row(r, str(y), {"B": "year", "C": "int", "D": "pct", "E": "pct", "F": "pct", "G": "pct", "H": "pct",
                            "I": "pct"}, check_label=False,
                B=y, C=int(m.sum()), D=xA[m].mean(), E=float(np.median(xA[m])), F=beat(xA[m]),
                G=xB[m].mean(), H=float(np.median(xB[m])), I=beat(xB[m]), J=f"{y} (n={int(m.sum())})")
    rep.row(77 + (y1 - y0) + 1, "All years", {"C": "int", "D": "pct", "E": "pct", "F": "pct", "G": "pct", "H": "pct",
                                              "I": "pct"}, C=len(tr), D=xA.mean(), E=float(np.median(xA)),
            F=beat(xA), G=xB.mean(), H=float(np.median(xB)), I=beat(xB))

    # 4c: V1 quarterly rebalance additions only
    q = robust_block(rep, 138, "4c. V1: quarterly-rebalance additions only",
                     {"C": ("V1 quarterly, Window A", xA, Q, d), "D": ("V1 quarterly, Window B", xB, Q, d)},
                     {"C": vA, "D": vB})
    # 4d: the other variants
    rep.text("B164", f"V4 Without {EXCL_YEAR_1} and {EXCL_YEAR_2}")
    rep.panel("4d. V3 winsorising cut-offs", {"C": "Window A", "D": "Window B"})
    rep.row(170, f"Lower cut-off: {text_pct(WINSOR_LO)} percentile", "pct", C=loA, D=loB)
    rep.row(171, f"Upper cut-off: {text_pct(WINSOR_HI)} percentile", "pct", C=hiA, D=hiB)
    v4 = f"V4 without {EXCL_YEAR_1} and {EXCL_YEAR_2}"
    heads = {"C": "V2 off-cycle", "D": "V3 winsorised", "E": v4, "F": "V5 sell the day before", "G": "V6 trimmed"}
    rep.text("E174", v4)
    rep.text("E197", v4)
    vA_ = {c: vA for c in "CDEFG"}
    vB_ = {c: vB for c in "CDEFG"}
    dA = robust_block(rep, 175, "4d. Variants, Window A",
                      {"C": (heads["C"], xA, OFF, d), "D": (heads["D"], wA, ALL, d), "E": (heads["E"], xA, V4, d),
                       "F": (heads["F"], xV5, ALL, d), "G": (heads["G"], xA, V6A, d)}, vA_)
    dB = robust_block(rep, 198, "4d. Variants, Window B",
                      {"C": (heads["C"], xB, OFF, d), "D": (heads["D"], wB, ALL, d), "E": (heads["E"], xB, V4, d),
                       "F": (heads["F"], xV5, ALL, d), "G": (heads["G"], xB, V6B, d)}, vB_)

    # Putting it together
    same = lambda b, v: yes(verdict(b) == v)
    first = {"C": same(halves["C"], vA), "D": same(halves["E"], vB)}
    second = {"C": same(halves["D"], vA), "D": same(halves["F"], vB)}
    result = {c: ("Robust" if first[c] == second[c] == "Yes" else
                  "Partly robust" if "Yes" in (first[c], second[c]) else "Not robust") for c in "CD"}
    others = {"C": sum(verdict(b) == vA for b in dA.values()), "D": sum(verdict(b) == vB for b in dB.values())}
    rep.panel("Putting it together", {"C": "Window A", "D": "Window B"})
    rep.row(223, "First half gives Test 1's verdict?", "", **first)
    rep.row(224, "Second half gives Test 1's verdict?", "", **second)
    rep.row(225, "Result for this window", "", **result)
    rep.row(226, "For reference: V1 quarterly only, verdict", "", C=verdict(q["C"]), D=verdict(q["D"]))
    rep.row(227, f"For reference: other variants giving Test 1's verdict (of {len(dA)})", "int", **others)
    rep.panel("Answer", {"C": "Result"})
    overall = ("Robust" if result["C"] == result["D"] == "Robust" else
               "Not robust" if result["C"] == result["D"] == "Not robust" else "Partly robust")
    rep.row(11, "Test 4 result", "", C=overall)
    rep.panels.insert(0, rep.panels.pop())
    # the subsets whose average is significantly above / below zero (used by the Summary sheet)
    sig = []
    sig += [(h[c], sig_word(halves[c].pG, halves[c].meanG)) for c in "CDEF"]
    sig += [(f"V1 quarterly, Window {w}", sig_word(q[c].pG, q[c].meanG)) for c, w in [("C", "A"), ("D", "B")]]
    sig += [(heads[c] + ("" if c == "F" else ", Window A"), sig_word(dA[c].pG, dA[c].meanG)) for c in "CDEFG"]
    sig += [(heads[c] + ", Window B", sig_word(dB[c].pG, dB[c].meanG)) for c in "CDEG"]
    return rep, result, sig


# ---------------------------------------------------------------------------------------------------------------------
# 2.12 Summary
# ---------------------------------------------------------------------------------------------------------------------
def report_summary(t1, t2, t3, t4, sig):
    rep = Report("Summary", "the four tests in one place")
    rep.panel("Results by test", {"D": "Window A", "E": "Window B"})
    rep.row(12, "Test 1: per trade", "", D=t1["D"], E=t1["E"])
    rep.row(13, "Test 2: risk-adjusted", "", D=t2["D"], E=t2["E"])
    rep.row(14, "Test 3: as an account", "", D=t3["D"], E=t3["E"])
    rep.row(15, "Test 4: robustness", "", D=t4["C"], E=t4["D"])
    passes = {c: yes(t1[c] == t2[c] == t3[c] == "Works" and t4[k] == "Robust") for c, k in [("D", "C"), ("E", "D")]}
    rep.row(16, "Passes all four tests?", "", **passes)
    def lst(word):
        """The Test 4 subsets whose average is significant in that direction, as the sheet lists them."""
        names = [lab for lab, s in sig if s == word]
        return "; ".join(names) + "." if names else "None."

    rep.panel("Test 4 subsets with a significant average (clustered t-test)", {"C": "Subsets"})
    rep.row(20, "Significantly above zero", "", C=lst("Yes, above zero"))
    rep.row(21, "Significantly below zero", "", C=lst("Yes, below zero"))
    rep.panel("Verdict", {"D": "Decision"})
    rep.row(8, "Verdict", "", D="Invest" if "Yes" in passes.values() else "Do not invest")
    rep.panels.insert(0, rep.panels.pop())
    return rep


# ---------------------------------------------------------------------------------------------------------------------
# 2.13 Extra Checks sheet, E1-E5 (its date averages and ranks are in Extra Checks Workings)
# ---------------------------------------------------------------------------------------------------------------------
def report_extra(tr, da, regs):
    rep = Report("Extra Checks", "E1 growth by half, E2 buying at the announcement close, E3 from 2020, "
                                 "E4 break test, E5 quarterly additions")
    T = len(da)
    d, year, names = tr.announced.values, tr.year.values, tr.name.values
    xA, xB = tr.AR_A.values, tr.AR_B.values
    H1, H2 = year <= SPLIT_YEAR, year > SPLIT_YEAR
    y0, y1 = int(year.min()), int(year.max())
    keys = {"C": "A", "D": "B", "E": "B1"}

    # E1: growth of the accounts by half
    h = regs["split_day"]
    v = {c: da["value_" + k].values for c, k in keys.items()}
    g1 = {c: v[c][h - 1] ** (DAYS_PER_YEAR / h) - 1 for c in keys}
    g2 = {c: (v[c][-1] / v[c][h - 1]) ** (DAYS_PER_YEAR / (T - h)) - 1 for c in keys}
    cols = {"C": "Long + SPY, A", "D": "Long + SPY, B", "E": "B1 SPY"}
    rep.panel("E1. Growth by half (value of $1 from the Daily Account)", cols)
    split_date = to_date(da.date.iat[h - 1])
    rep.row(17, f"First half: last day ({text_dmy(split_date)})", "", **{c: split_date for c in keys})
    rep.row(18, "First half: trading days", "int", **{c: h for c in keys})
    rep.row(19, "First half: value of $1 at the end ($1 at the 31 Dec 2009 close)", "usd",
            **{c: v[c][h - 1] for c in keys})
    rep.row(20, f"First half: growth a year = value^({DAYS_PER_YEAR} ÷ days) − 1", "pct", **g1)
    rep.row(21, "Second half: trading days", "int", **{c: T - h for c in keys})
    rep.row(22, "Second half: value of $1 at the end", "usd", **{c: v[c][-1] for c in keys})
    rep.row(23, f"Second half: growth a year = (end ÷ start)^({DAYS_PER_YEAR} ÷ days) − 1", "pct", **g2)
    rep.row(24, "First half: growth a year above B1", "pct", C=g1["C"] - g1["E"], D=g1["D"] - g1["E"])
    rep.row(25, "Second half: growth a year above B1", "pct", C=g2["C"] - g2["E"], D=g2["D"] - g2["E"])
    rA, rB = regs["capm_A_h2"], regs["capm_B_h2"]
    rep.row(26, "Second half: CAPM alpha a year (Test 4a)", "pct", C=rA.alpha_year, D=rB.alpha_year)
    rep.row(27, "…p-value (Newey–West)", "p", C=t_dist_2t(rA.t, rA.T - 2), D=t_dist_2t(rB.t, rB.T - 2))

    # E2: bought at the announcement close (not tradable: an upper bound)
    full_rows = [(0, "N: trades", "N", "int"),
                 (1, "Mean abnormal return (each trade once)", "mean", "pct"),
                 (2, "Median", "median", "pct"),
                 (3, "% of trades that beat SPY (k ÷ non-zero trades)", "beat", "pct"),
                 (4, "Announcement dates (G)", "G", "int"),
                 (5, "Mean of the date averages (each date once)", "meanG", "pct"),
                 (6, "Clustered SE", "seG", "pct"),
                 (7, "Clustered t = mean of date averages ÷ SE", "tG", "num"),
                 (8, "p-value: T.DIST.2T(|t|, G − 1)", "pG", "p")]
    wil_rows = [(0, "Wilcoxon n (non-zero trades)", "n", "int"),
                (1, "W+: sum of ranks of positive trades", "W", "num2"),
                (2, "z = (W+ − n(n + 1) ÷ 4) ÷ √[n(n + 1)(2n + 1) ÷ 24]", "z", "num"),
                (3, "Wilcoxon p-value", "pW", "p")]
    e2 = {c: stat_block(tr[k].values, d)
          for c, k in [("C", "CAR_inc_open"), ("D", "CAR_inc_close"), ("E", "CAR_buy_open")]}
    rep.panel("E2. Bought at the announcement close (CAR since the announcement close; all trades)",
              {"C": "To inclusion open", "D": "To inclusion close", "E": "Overnight (Ann+1 open)"})
    for j, label, f, fmt in full_rows:
        rep.row(34 + j, label, fmt, **{c: getattr(b, f) for c, b in e2.items()})
    for j, label, f, fmt in wil_rows:
        rep.row(43 + j, label, fmt, **{c: getattr(b, f) for c, b in e2.items()})
    rep.row(47, "Sign test p-value: 2 × BINOM.DIST(min(k, n − k), n, 0.5)", "p", **{c: b.pS for c, b in e2.items()})
    hh = {"C": f"To the inclusion open, {y0}–{SPLIT_YEAR}", "D": f"To the inclusion open, {SPLIT_YEAR + 1}–{y1}",
          "E": f"To the inclusion close, {y0}–{SPLIT_YEAR}", "F": f"To the inclusion close, {SPLIT_YEAR + 1}–{y1}"}
    for c, lab in hh.items():
        rep.text(f"{c}50", lab)
    by_half = [("C", "CAR_inc_open", H1), ("D", "CAR_inc_open", H2),
               ("E", "CAR_inc_close", H1), ("F", "CAR_inc_close", H2)]
    e2h = {c: stat_block(tr[k].values[m], d[m]) for c, k, m in by_half}
    rep.panel("E2. By half", hh)
    for r, (j, label, f, fmt) in zip(range(51, 57), [full_rows[0]] + full_rows[4:9]):
        rep.row(r, label, fmt, **{c: getattr(b, f) for c, b in e2h.items()})

    # E3: from START_2020 onwards
    F20 = d >= np.datetime64(START_2020)
    F20XA, F20XB = F20 & (xA < xA[F20].max()), F20 & (xB < xB[F20].max())
    e3 = {c: stat_block(x[m], d[m]) for c, x, m in [("C", xA, F20), ("D", xA, F20XA), ("E", xB, F20), ("F", xB, F20XB)]}
    largest = lambda x, m, k: names[int(np.nonzero(x == np.sort(x[m])[::-1][k - 1])[0][0])]
    rep.panel(f"E3. Trades announced from {text_dmy(START_2020)}",
              {"C": "Window A", "D": "A without largest", "E": "Window B", "F": "B without largest"})
    e3_fields = [("N", "int"), ("mean", "pct"), ("median", "pct"), ("beat", "pct"), ("G", "int"), ("meanG", "pct"),
                 ("seG", "pct"), ("tG", "num"), ("pG", "p"), ("lo", "pct"), ("n", "int"), ("W", "num2"), ("z", "num"),
                 ("pW", "p"), ("mde", "pct")]
    e3_labels = ["N: trades",
                 "Mean abnormal return (each trade once)",
                 "Median",
                 "% of trades that beat SPY (k ÷ non-zero trades)",
                 "Announcement dates (G)",
                 "Mean of the date averages (each date once)",
                 "Clustered SE",
                 "Clustered t = mean of date averages ÷ SE",
                 "p-value: T.DIST.2T(|t|, G − 1)",
                 f"Lower end of the {text_pct(1 - ALPHA)} interval: mean − t* × SE",
                 "Wilcoxon n (non-zero trades)",
                 "W+: sum of ranks of positive trades",
                 "z = (W+ − n(n + 1) ÷ 4) ÷ √[n(n + 1)(2n + 1) ÷ 24]",
                 "Wilcoxon p-value",
                 f"MDE = ({ZCRIT:.2f} + {ZPOW:.2f}) × clustered SE"]
    for j, ((f, fmt), label) in enumerate(zip(e3_fields, e3_labels)):
        rep.row(63 + j, label, fmt, **{c: getattr(b, f) for c, b in e3.items()})
    rep.row(78, "Verdict (Test 1's rule)", "", **{c: verdict(b) for c, b in e3.items()})
    rep.row(79, "Largest trade left out", "", D=largest(xA, F20, 1), F=largest(xB, F20, 1))
    # the account from the first trading day of 2020
    p20 = regs["p20"]                                       # day number of the last day before START_2020
    n20 = T - p20
    g = {c: (v[c][-1] / v[c][p20 - 1]) ** (DAYS_PER_YEAR / n20) - 1 for c in keys}
    srd = {c: sharpe_daily(da["x_" + k].values[p20:]) for c, k in keys.items()}
    jk = {c: jobson_korkie(da["x_" + keys[c]].values[p20:], da.x_B1.values[p20:]) for c in "CD"}
    c20 = {c: regs[f"capm_{keys[c]}_2020"] for c in "CD"}
    rep.panel(f"E3. The account from {text_dmy(START_2020)}", cols)
    rep.row(83, "Start: value taken on this day's close", "", **{c: to_date(da.date.iat[p20 - 1]) for c in keys})
    rep.row(84, "Trading days from the start", "int", **{c: n20 for c in keys})
    rep.row(85, "Value of $1 at the start", "usd", **{c: v[c][p20 - 1] for c in keys})
    rep.row(86, "Value of $1 at the end", "usd", **{c: v[c][-1] for c in keys})
    rep.row(87, f"Growth a year = (end ÷ start)^({DAYS_PER_YEAR} ÷ days) − 1", "pct", **g)
    rep.row(88, "Growth a year above B1", "pct", C=g["C"] - g["E"], D=g["D"] - g["E"])
    rep.row(89, "Sharpe ratio, daily = average excess return ÷ its st. dev.", "num", **srd)
    rep.row(90, f"Sharpe ratio, annualised = daily × √{DAYS_PER_YEAR}", "num",
            **{c: s * math.sqrt(DAYS_PER_YEAR) for c, s in srd.items()})
    rep.row(91, "ρ: correlation with B1's daily excess return", "num", **{c: jk[c].rho for c in "CD"})
    rep.row(92, "Jobson–Korkie z (Memmel, 2003), as in Test 3", "num", **{c: jk[c].z for c in "CD"})
    rep.row(93, "…p-value (two-sided)", "p", **{c: jk[c].p for c in "CD"})
    rep.row(94, "CAPM alpha a year (daily α × days a year)", "pct", **{c: c20[c].alpha_year for c in "CD"})
    rep.row(95, "…t-statistic (Newey–West)", "num", **{c: c20[c].t for c in "CD"})
    rep.row(96, "…p-value: T.DIST.2T(|t|, days − 2)", "p", **{c: t_dist_2t(c20[c].t, c20[c].T - 2) for c in "CD"})

    # E4: one break in the mean of the date averages (sup-F, Bai and Perron critical values)
    ALL = np.ones(len(tr), bool)
    XA, XB = xA < xA.max(), xB < xB.max()                   # every trade except the largest in that window
    e4 = {c: sup_f_break(x[m], d[m]) for c, x, m in [("C", xA, ALL), ("D", xB, ALL), ("E", xA, XA), ("F", xB, XB)]}
    rep.panel("E4. Break test on the date averages (one value per announcement date, in date order)",
              {"C": "Window A", "D": "Window B", "E": "A without largest", "F": "B without largest"})
    e4_rows = [(104, "Announcement dates (G)", "G", "int"),
               (105, "Dates kept out at each end = INT(share × G)", "lo", "int"),
               (106, "sup-F: the largest F", "sup", "num"),
               (107, "Break after: the last date of the first period", "last", ""),
               (108, "…the first date of the second period", "next", ""),
               (109, "Dates before the break", "k", "int"),
               (110, "Dates after the break", "k2", "int"),
               (111, "Mean of the date averages before", "m1", "pct"),
               (112, "Mean of the date averages after", "m2", "pct"),
               (113, "Critical value at 5% (Settings)", "cv5", "num2"),
               (114, "Critical value at 1% (Settings)", "cv1", "num2"),
               (115, "Significant break?", "sig", "")]
    for r, label, f, fmt in e4_rows:
        rep.row(r, label, fmt, **{c: getattr(b, f) for c, b in e4.items()})
    rep.row(116, "Largest trade left out", "", E=largest(xA, ALL, 1), F=largest(xB, ALL, 1))

    # E5: quarterly additions (V1), by half and without the N_DROP_V1 largest winners
    Q = tr.quarterly.values == "Yes"
    drop = lambda x: Q & np.array([(x[Q] > xi).sum() >= N_DROP_V1 for xi in x])
    QXA, QXB = drop(xA), drop(xB)
    e5 = {c: stat_block(x[m], d[m]) for c, x, m in
          [("C", xA, Q), ("D", xA, Q & H1), ("E", xA, Q & H2), ("F", xA, QXA),
           ("G", xB, Q), ("H", xB, Q & H1), ("I", xB, Q & H2), ("J", xB, QXB)]}
    hq = {"C": "A: all", "D": f"A: {y0}–{SPLIT_YEAR}", "E": f"A: {SPLIT_YEAR + 1}–{y1}",
          "F": f"A: without {N_DROP_V1}", "G": "B: all", "H": f"B: {y0}–{SPLIT_YEAR}",
          "I": f"B: {SPLIT_YEAR + 1}–{y1}", "J": f"B: without {N_DROP_V1}"}
    for c in "DEHI":
        rep.text(f"{c}121", hq[c])
    rep.panel("E5. Quarterly-rebalance additions (V1)", hq)
    e5_rows = [(122, "N: trades", "N", "int"), (123, "Announcement dates (G)", "G", "int"),
               (124, "Mean abnormal return (each trade once)", "mean", "pct"),
               (125, "Mean of the date averages (each date once)", "meanG", "pct"), (126, "Clustered SE", "seG", "pct"),
               (127, "Clustered t = mean of date averages ÷ SE", "tG", "num"),
               (128, "p-value: T.DIST.2T(|t|, G − 1)", "pG", "p"),
               (129, f"MDE = ({ZCRIT:.2f} + {ZPOW:.2f}) × clustered SE", "mde", "pct")]
    for r, label, f, fmt in e5_rows:
        rep.row(r, label, fmt, **{c: getattr(b, f) for c, b in e5.items()})
    for k in (1, 2, 3):
        rep.row(129 + k, f"Left out: largest winner {k}", "",
                F=largest(xA, Q, k) if k <= N_DROP_V1 else "–", J=largest(xB, Q, k) if k <= N_DROP_V1 else "–")
    return rep


def sup_f_break(x, dates):
    """Sup-F test for one break in the mean of the date averages (Andrews 1993; Bai and Perron 1998).
    y = date averages in date order (G values). For k = LO..G-LO-1 dates in the first segment:
    SSR_k = sum of squares around each segment's mean, F_k = (SSR_0 - SSR_k) / (SSR_k / (G - 2))."""
    by_date = pd.Series(x).groupby(dates).mean().sort_index()
    y, G = by_date.values, len(by_date)
    lo = int(math.floor(BREAK_TRIM * G))
    cs = np.cumsum(y)
    q, SG = (y * y).sum(), cs[-1]
    ssr0 = q - SG ** 2 / G
    ks = np.arange(lo, G - lo)
    s = cs[ks - 1]
    ssr = q - s ** 2 / ks - (SG - s) ** 2 / (G - ks)
    F = (ssr0 - ssr) / (ssr / (G - 2))
    j = int(np.argmax(F))                                   # first k with the largest F
    k = int(ks[j])
    sup = float(F[j])
    sig = "Yes, at 1%" if sup > BAI_PERRON_CV_1 else "Yes, at 5%" if sup > BAI_PERRON_CV_5 else "No"
    return SimpleNamespace(G=G, lo=lo, sup=sup, last=to_date(by_date.index[k - 1]), next=to_date(by_date.index[k]),
                           k=k, k2=G - k, m1=s[j] / k, m2=(SG - s[j]) / (G - k), cv5=BAI_PERRON_CV_5,
                           cv1=BAI_PERRON_CV_1, sig=sig)

# ---------------------------------------------------------------------------------------------------------------------
# 2.14 Run every calculation
# ---------------------------------------------------------------------------------------------------------------------
def all_regressions(da):
    """The regressions of the Regression Workings sheet (day numbers are 1-based positions in the Daily Account)."""
    z = day_number(da.date, da.attrs["ff_last"])                         # last day with factors (4,190 in v12)
    h = day_number(da.date, dt.date(SPLIT_YEAR, 12, 31))                 # last day of the first half (2,013)
    p20 = day_number(da.date, START_2020 - dt.timedelta(days=1))         # last day before START_2020 (2,516)
    regs = {"full_T": z, "split_day": h, "p20": p20}
    for k in ["A", "B"]:
        y = "x_" + k
        regs[f"capm_{k}"] = regression(da, y, 1, z)                      # regressions 1, 2
        regs[f"ff4_{k}"] = regression(da, y, 1, z, four_factor=True)     # regressions 3, 4
        regs[f"capm_{k}_h1"] = regression(da, y, 1, h)                   # regressions 5, 7
        regs[f"capm_{k}_h2"] = regression(da, y, h + 1, z)               # regressions 6, 8
        regs[f"capm_{k}_2020"] = regression(da, y, p20 + 1, z)           # Extra Checks E3
    return regs



def run_calculations(raw):
    """PART 2: every calculation, in the order of the workbook's sheets."""
    spy = spy_close_by_date(raw)
    beta = beta_estimation(raw)
    tr = trade_returns(raw, beta)
    hd = holding_days(raw, tr)
    da = daily_account(raw, hd, spy)
    regs = all_regressions(da)
    settings = report_settings(da)
    t1, t1a, t1b = report_test1(tr)
    t2 = report_test2(tr, beta, t1a, t1b)
    t3, t3res = report_test3(tr, hd, da, regs)
    t4, t4res, sig = report_test4(tr, da, regs, t1a, t1b)
    summary = report_summary({"D": t1.cells["D141"], "E": t1.cells["E141"]},
                             {"D": t2.cells["D122"], "E": t2.cells["E122"]}, t3res, t4res, sig)
    extra = report_extra(tr, da, regs)
    start_close = to_date(spy.index[spy.index < da.date.iat[0]][-1])      # the portfolio's $1 goes in at this close
    return SimpleNamespace(raw=raw, beta=beta, tr=tr, hd=hd, da=da, regs=regs, start_close=start_close,
                           reports=[settings, summary, t1, t2, t3, t4, extra])


# =====================================================================================================================
# PART 3. THE PAPER'S TABLES AND FIGURES, printed from the results of Part 2
# =====================================================================================================================
def excel_round(x, d):
    """x rounded to d decimals the way Excel displays it: 15 significant digits, then half away from zero."""
    q = Decimal(format(float(x), ".15g")).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    return q if q != 0 else abs(q)


def pc(x, d=2, sign=False):
    """A decimal as a percentage, e.g. -0.0064 -> -0.64%; sign=True writes + before a gain."""
    q = excel_round(Decimal(format(float(x), ".15g")) * 100, d)
    return ("+" if sign and q > 0 else "") + f"{q:.{d}f}%"


def nm(x, d=2, sign=False):
    q = excel_round(x, d)
    return ("+" if sign and q > 0 else "") + f"{q:.{d}f}"


def pv(p):
    """A p-value: three decimals, or <0.001."""
    return "<0.001" if p < 0.001 else nm(p, 3)


def peq(p):
    """'p = 0.077' or 'p < 0.001', for sentences."""
    return "p < 0.001" if p < 0.001 else f"p = {nm(p, 3)}"


def star(p):
    """* = significant at the 5% level."""
    return "*" if p < ALPHA else ""


def usd(x):
    return f"${excel_round(x, 2):.2f}"


def dmy(d):
    return d if isinstance(d, str) else text_dmy(d)


def in_notebook():
    try:
        from IPython import get_ipython
        return get_ipython() is not None and "IPKernelApp" in get_ipython().config
    except Exception:
        return False


def show(title, header, rows, notes=()):
    """One table: a formatted table in a notebook (Colab, Jupyter), aligned text otherwise.
    rows = [(label, cell, cell, ...)], all cells as text."""
    print("\n" + title)
    if in_notebook():
        from IPython.display import HTML, display
        df = pd.DataFrame([r[1:] for r in rows], index=[r[0] for r in rows], columns=header)
        try:                                         # labels on the left, numbers on the right
            display(HTML(df.style.set_table_styles([
                {"selector": "th.row_heading", "props": [("text-align", "left")]},
                {"selector": "th.col_heading", "props": [("text-align", "right")]},
                {"selector": "td", "props": [("text-align", "right"), ("padding", "2px 10px")]}]).to_html()))
        except Exception:                            # (styling needs the jinja2 package)
            display(df)
    else:
        lw = max(len(r[0]) for r in rows) + 2
        cw = [max(len(h), *(len(r[j + 1]) for r in rows)) + 2 for j, h in enumerate(header)]
        line = "-" * (lw + sum(cw))
        print(line)
        print(" " * lw + "".join(h.rjust(w) for h, w in zip(header, cw)))
        print(line)
        for r in rows:
            print(r[0].ljust(lw) + "".join(c.rjust(w) for c, w in zip(r[1:], cw)))
        print(line)
    for note in notes:
        print(note)


def print_paper_tables(res):
    c = {rep.sheet: rep.cells for rep in res.reports}
    t1, t2, t3, t4 = c["Test 1 Per Trade"], c["Test 2 Risk-Adjusted"], c["Test 3 Portfolio"], c["Test 4 Robustness"]
    ex, sm = c["Extra Checks"], c["Summary"]
    tr, da, beta, info = res.tr, res.da, res.beta, res.raw.info
    years = tr.year.values
    y0, y1 = int(years.min()), int(years.max())
    h1, h2 = f"{y0}-{str(SPLIT_YEAR)[2:]}", f"{SPLIT_YEAR + 1}-{str(y1)[2:]}"
    first, last = to_date(da.date.iat[0]), to_date(da.date.iat[-1])
    hday = res.regs["split_day"]
    ff_last = to_date(da.attrs["ff_last"])
    star_note = "* Significant at the 5% level (two-sided)."

    print("\n" + "=" * 100)
    print("PART 3. THE PAPER'S TABLES AND FIGURES")
    print("=" * 100)
    print(f"\nVerdict: {sm['D8']}")
    print("\nThe sample and the data")
    print(f"  Additions studied: {len(tr)}, announced {dmy(tr.announced.min())} to {dmy(tr.announced.max())}, on "
          f"{tr.announced.nunique()} announcement dates; {(tr.quarterly == 'Yes').sum()} at S&P's quarterly "
          f"rebalance, {(tr.quarterly == 'No').sum()} off-cycle")
    print(f"  Trading days from the announcement to inclusion: median {tr.days.median():.0f} "
          f"(range {tr.days.min()} to {tr.days.max()})")
    print(f"  Prices: {info.crsp} additions from CRSP, {info.yahoo} from Yahoo Finance; SPY from CRSP to "
          f"{dmy(CRSP_END)}, then Yahoo Finance")
    print(f"  Holding periods with a stock split: {info.splits}; with an ex-dividend date: {info.n_div} (largest "
          f"dividend {pc(info.max_div, 1)} of the price; returns leave dividends out)")
    few = beta.days.values < MIN_BETA_DAYS
    print(f"  Betas (trading days -{BETA_FIRST_DAY} to -{BETA_LAST_DAY}): {(~few).sum()} estimated, {few.sum()} set to "
          f"1 (fewer than {MIN_BETA_DAYS} days); average beta {nm(beta.beta.mean())}")
    print(f"  Portfolio: {len(da):,} trading days, from the {dmy(res.start_close)} close to {dmy(last)}; "
          f"Fama-French factors to {dmy(ff_last)}")
    ch = info.changed
    moved = ", ".join(ch[:-1]) + " and " + ch[-1] if len(ch) > 1 else "".join(ch)
    print("  Data check: " + (f"the downloaded {moved} differ from the data behind the paper, so a number below can "
                              f"differ slightly from the paper" if moved else
                              "the downloaded data are the same as the data behind the paper"))

    # ---------------------------------------------------------------------------------------------- Table 1
    verdict_word = lambda col: "Invest" if sm[f"{col}16"] == "Yes" else "Do not invest"
    show(f"Table 1. Summary of results: the strategy against SPY, {first:%B %Y} to {last:%B %Y}",
         ["Strategy, Window A", "Strategy, Window B", "Benchmark"],
         [("Median trade: return above SPY", pc(t1["D72"]) + star(t1["D78"]), pc(t1["E72"]) + star(t1["E78"]), "-"),
          ("Median trade, adjusted for beta", pc(t2["D84"]) + star(t2["D90"]), pc(t2["E84"]) + star(t2["E90"]), "-"),
          ("Average trade: return above SPY", pc(t1["D52"]) + star(t1["D56"]), pc(t1["E52"]) + star(t1["E56"]), "-"),
          ("Average trade, adjusted for beta", pc(t2["D64"]) + star(t2["D68"]), pc(t2["E64"]) + star(t2["E68"]), "-"),
          ("Trades that beat SPY", pc(t1["D86"], 1) + star(t1["D87"]), pc(t1["E86"], 1) + star(t1["E87"]), "-"),
          ("Trades that beat SPY, adjusted for beta", pc(t2["D98"], 1) + star(t2["D99"]),
           pc(t2["E98"], 1) + star(t2["E99"]), "-"),
          ("Portfolio: annualised return", pc(t3["D37"]), pc(t3["D79"]), pc(t3["E37"])),
          ("Portfolio: Sharpe ratio", nm(t3["D44"]), nm(t3["D86"]), nm(t3["E44"])),
          ("Verdict", verdict_word("D"), verdict_word("E"), "")],
         ["* Significantly different at the 5% level from zero (median, average) or from half of all trades.",
          "The average trade counts each announcement date once."])

    # ---------------------------------------------------------------------------------------------- Table 3
    cols = [(t1, "D", 0, "D141"), (t1, "E", 0, "E141"), (t2, "D", 12, "D122"), (t2, "E", 12, "E122")]
    cell = lambda s, col, off, r: s[f"{col}{r + off}"]
    show("Table 3. Results per trade: the added stock's return minus SPY's over the same hours",
         ["Window A", "Window B", "A, beta-adjusted", "B, beta-adjusted"],
         [("Median trade (p)", *[f"{pc(cell(s, k, o, 72))} ({pv(cell(s, k, o, 78))})" for s, k, o, _ in cols]),
          ("Average trade (p)", *[f"{pc(cell(s, k, o, 52))} ({pv(cell(s, k, o, 56))})" for s, k, o, _ in cols]),
          ("Trades that beat SPY (p)",
           *[f"{pc(cell(s, k, o, 86), 1)} ({pv(cell(s, k, o, 87))})" for s, k, o, _ in cols]),
          ("Smallest average gain the test could detect", *[pc(cell(s, k, o, 94)) for s, k, o, _ in cols]),
          ("Result", *[s[v] for s, _, _, v in cols])],
         ["p-values: Wilcoxon signed-rank test (median), t-test with each announcement date counted once (average),",
          "sign test (trades that beat SPY). In the beta-adjusted columns a trade beats SPY if it beats beta x SPY."])

    # ---------------------------------------------------------------------------------------------- Figure 1
    moments = [(103, "Announcement close"), (104, "Ann+1 open (the strategy buys)"), (105, "Ann+1 close"),
               (106, "Close the day before inclusion"), (107, "Inclusion open (sells, Window A)"),
               (108, "Inclusion close (sells, Window B)")]
    show("Figure 1 (data, Table B2). Return above SPY since the announcement close",
         ["Average (each date once)", "95% range", "Average (each trade once)"],
         [(lab, pc(t1[f"D{r}"], sign=True),
           "-" if t1[f"E{r}"] == 0 else f"{pc(t1[f'F{r}'], sign=True)} to {pc(t1[f'G{r}'], sign=True)}",
           pc(t1[f"H{r}"], sign=True)) for r, lab in moments],
         ["Return above SPY = the stock's price relative to its announcement close minus SPY's price relative to its "
          "announcement close."])

    # ---------------------------------------------------------------------------------------------- Table 4
    A, B = 0, 42                                     # Panel A and Panel B rows of the Test 3 sheet
    p = lambda r, off: t3[f"D{r + off}"]
    show(f"Table 4. The Strategy Portfolio against the Benchmark Portfolio, {dmy(first)} to {dmy(last)}",
         ["Strategy, Window A", "Strategy, Window B", "Benchmark"],
         [("Value of $1 at the end", usd(p(39, A)), usd(p(39, B)), usd(t3["E39"])),
          ("Annualised return", pc(p(37, A)), pc(p(37, B)), pc(t3["E37"])),
          ("Annualised return above Benchmark", pc(p(38, A), sign=True), pc(p(38, B), sign=True), "-"),
          ("Volatility a year", pc(p(40, A)), pc(p(40, B)), pc(t3["E40"])),
          ("Sharpe ratio (p against Benchmark)", f"{nm(p(44, A))} ({pv(p(49, A))})", f"{nm(p(44, B))} ({pv(p(49, B))})",
           nm(t3["E44"])),
          ("CAPM alpha per year (t)", f"{pc(p(55, A), sign=True)} ({nm(p(56, A))})",
           f"{pc(p(55, B), sign=True)} ({nm(p(56, B))})", "-"),
          ("Four-factor alpha per year (t)", f"{pc(p(59, A), sign=True)} ({nm(p(60, A))})",
           f"{pc(p(59, B), sign=True)} ({nm(p(60, B))})", "-"),
          ("Market beta", nm(p(58, A)), nm(p(58, B)), "-"),
          ("Maximum drawdown", pc(p(68, A), 1), pc(p(68, B), 1), pc(t3["E68"], 1)),
          ("Drawdown: peak before the fall", dmy(p(69, A)), dmy(p(69, B)), dmy(t3["E69"])),
          ("Drawdown: trough", dmy(p(70, A)), dmy(p(70, B)), dmy(t3["E70"])),
          ("Drawdown: back at the peak by", dmy(p(71, A)), dmy(p(71, B)), dmy(t3["E71"])),
          ("Days holding added stocks", pc(p(35, A), 1), pc(p(35, B), 1), "-"),
          ("Result", t3["D123"], t3["E123"], "-")],
         [f"Alphas run to {dmy(ff_last)} (the last day of French's data), with Newey-West standard errors "
          f"({NW_LAGS} lags). Their p-values: CAPM {pv(p(57, A))} (A) and {pv(p(57, B))} (B); four-factor "
          f"{pv(p(61, A))} (A) and {pv(p(61, B))} (B).",
          "Sharpe ratio p-values: Jobson and Korkie's test with Memmel's correction."])

    # ---------------------------------------------------------------------------------------------- Figure 2
    v = da.set_index("date")[["value_A", "value_B", "value_B1"]]
    ends = v.groupby(v.index.year).tail(1)
    rows = [(f"{dmy(res.start_close)} (start)", "$1.00", "$1.00", "$1.00")]
    rows += [(dmy(d), usd(r.value_A), usd(r.value_B), usd(r.value_B1)) for d, r in ends.iterrows()]
    show(f"Figure 2 (data). Value of $1 invested at the close on {dmy(res.start_close)}, at the end of each year",
         ["Strategy, Window A", "Strategy, Window B", "Benchmark"], rows,
         ["Every trading day's value is in daily_account.csv (option --save)."])

    # ---------------------------------------------------------------------------------------------- Table 5
    hc = ["C", "D", "E", "F"]                         # A first half, A second half, B first half, B second half
    acct = [("C", "61", "C", "64"), ("D", "61", "D", "64"), ("C", "62", "C", "68"), ("D", "62", "D", "68")]
    show("Table 5. The two halves of the sample",
         [f"Window A, {h1}", f"Window A, {h2}", f"Window B, {h1}", f"Window B, {h2}"],
         [("Trades (announcement dates)", *[f"{t4[k + '18']} ({t4[k + '22']})" for k in hc]),
          ("Median trade (p)", *[f"{pc(t4[k + '20'], sign=True)} ({pv(t4[k + '32'])})" for k in hc]),
          ("Average trade (p)", *[f"{pc(t4[k + '23'], sign=True)} ({pv(t4[k + '26'])})" for k in hc]),
          ("Trades that beat SPY", *[pc(t4[k + "21"], 1) for k in hc]),
          ("Smallest average gain the test could detect", *[pc(t4[k + "33"]) for k in hc]),
          ("Strategy Portfolio Sharpe ratio (Benchmark)",
           *[f"{nm(t4[k + r])} ({nm(t4[k + '63'])})" for k, r, _, _ in acct]),
          ("CAPM alpha a year (p)",
           *[f"{pc(t4[k + r], sign=True)} ({pv(t4[k + str(int(r) + 2)])})" for _, _, k, r in acct]),
          ("Result", *[t4[k + "35"] for k in hc])],
         [f"Trades are split by the year the addition was announced; the portfolio rows split the trading days at "
          f"{dmy(to_date(da.date.iat[hday - 1]))}.",
          f"Alphas run to {dmy(ff_last)}, with Newey-West standard errors. Inconclusive means no significant result "
          f"and a test too weak to detect a gain of {pc(MIN_EFFECT, 0)} per trade.",
          f"Do the halves differ? Welch t-test {peq(t4['C50'])} (Window A) and {peq(t4['D50'])} (Window B); "
          f"Mann-Whitney {peq(t4['C55'])} and {peq(t4['D55'])}.",
          f"Test 4 result: {t4['C11']} (first half gives the full sample's answer: {t4['C223']} (A), {t4['D223']} "
          f"(B); second half: {t4['C224']} (A), {t4['D224']} (B))."])

    # ---------------------------------------------------------------------------------------------- Figure 3
    rows = []
    for y in range(y0, y1 + 1):
        r = 77 + (y - y0)
        rows.append((str(y), str(t4[f"C{r}"]), pc(t4[f"D{r}"], sign=True), pc(t4[f"G{r}"], sign=True)))
    r = 77 + (y1 - y0) + 1
    rows.append(("All years", str(t4[f"C{r}"]), pc(t4[f"D{r}"], sign=True), pc(t4[f"G{r}"], sign=True)))
    show("Figure 3 (data). The average trade in each year (return above SPY, each trade once)",
         ["Trades", "Window A", "Window B"], rows)

    # ---------------------------------------------------------------------------------------------- Table B1
    def b1(s, k, o, v):
        g = lambda r: s[f"{k}{r + o}"]
        return [str(g(36)), str(g(50)), f"{pc(g(37))} ({pv(g(41))})", f"{pc(g(52))}{star(g(56))} ({pv(g(56))})",
                pc(g(54)), nm(g(55)), f"{pc(g(63))} to {pc(g(64), sign=True)}", pc(g(94)),
                f"{pc(g(72))}{star(g(78))} ({pv(g(78))})", f"{pc(g(86), 1)}{star(g(87))} ({pv(g(87))})", s[v]]
    labels = ["Trades", "Announcement dates", "Average, each trade counted once (p)",
              "Average, each date counted once (p)", "  standard error", "  t", "  95% range for the date average",
              "Smallest average gain the test could detect", "Median (p)", "Trades that beat SPY (p)", "Result"]
    cells = [b1(*x) for x in cols]
    show("Table B1. Results per trade in full: the added stock's return minus SPY's over the same hours",
         ["Window A", "Window B", "A, beta-adjusted", "B, beta-adjusted"],
         [(lab, *[cc[j] for cc in cells]) for j, lab in enumerate(labels)],
         [f"Betas from each stock's daily returns on trading days -{BETA_FIRST_DAY} to -{BETA_LAST_DAY} before the "
          f"announcement (1 for stocks with fewer than {MIN_BETA_DAYS} days). " + star_note])

    # ---------------------------------------------------------------------------------------------- Figure B1
    def bins(xs):
        cnt = {k: 0 for k in range(-20, 20)}
        for x in xs:
            v_ = x * 100
            cnt[-20 if v_ < -20 else 19 if v_ >= 20 else int(v_ // 1)] += 1
        return cnt
    bA, bB = bins(tr.AR_A.values), bins(tr.AR_B.values)
    sg = lambda k: f"{k:+d}%" if k else "0%"
    used = [k for k in range(-20, 20) if bA[k] or bB[k]]
    best = int(np.argmax(tr.AR_A.values))
    show("Figure B1 (data). The spread of returns above SPY across the trades: number of trades in each "
         "1-percentage-point bin",
         ["Window A", "Window B"],
         [(f"{sg(k)} to {sg(k + 1)}" if -20 < k < 19 else ("below -19%" if k == -20 else "+19% or more"),
           str(bA[k]), str(bB[k])) for k in range(used[0], used[-1] + 1)],
         [f"Median trade: {pc(np.median(tr.AR_A.values))} (A) and {pc(np.median(tr.AR_B.values))} (B). Largest "
          f"winner: {tr.label.iat[best]}, {pc(tr.AR_A.iat[best], 1, sign=True)} in Window A."])

    # ---------------------------------------------------------------------------------------------- Table C1
    labels = [("Trades", lambda k: str(t4[k + "18"])), ("Announcement dates", lambda k: str(t4[k + "22"])),
              ("Average, each trade counted once", lambda k: pc(t4[k + "19"], sign=True)),
              ("Average, each date counted once", lambda k: pc(t4[k + "23"], sign=True) + star(t4[k + "26"])),
              ("  standard error", lambda k: pc(t4[k + "24"])), ("  t", lambda k: nm(t4[k + "25"], sign=True)),
              ("  p", lambda k: pv(t4[k + "26"])),
              ("Lower end of the 95% range", lambda k: pc(t4[k + "27"], sign=True)),
              ("Smallest average gain the test could detect", lambda k: pc(t4[k + "33"])),
              ("Median", lambda k: pc(t4[k + "20"], sign=True) + star(t4[k + "32"])),
              ("  p (Wilcoxon)", lambda k: pv(t4[k + "32"])),
              ("Trades that beat SPY", lambda k: pc(t4[k + "21"], 1)), ("Result", lambda k: t4[k + "35"])]
    show("Table C1. The two halves trade by trade, in full",
         [f"Window A, {h1}", f"Window A, {h2}", f"Window B, {h1}", f"Window B, {h2}"],
         [(lab, *[f(k) for k in hc]) for lab, f in labels],
         [f"Trades are split by the year the addition was announced (first half to 31 December {SPLIT_YEAR}). "
          + star_note])

    # ---------------------------------------------------------------------------------------------- Table C2
    def c2(n_a, n_b, mA, pA, aA, gA, mB, pB, aB, gB, vA, vB):
        return [str(n_a) if n_a == n_b else f"{n_a} / {n_b}",
                f"{pc(mA, sign=True)}{star(pA)} ({pv(pA)})", f"{pc(aA, sign=True)}{star(gA)} ({pv(gA)})",
                f"{pc(mB, sign=True)}{star(pB)} ({pv(pB)})", f"{pc(aB, sign=True)}{star(gB)} ({pv(gB)})",
                vA if vA == vB else f"A: {vA}; B: {vB}"]
    rows = [("Full sample (Table 3)", *c2(t1["D36"], t1["E36"], t1["D72"], t1["D78"], t1["D52"], t1["D56"],
                                         t1["E72"], t1["E78"], t1["E52"], t1["E56"], t1["D141"], t1["E141"])),
            ("V1 Quarterly reviews only", *c2(t4["C138"], t4["D138"], t4["C140"], t4["C152"], t4["C143"],
                                             t4["C146"], t4["D140"], t4["D152"], t4["D143"], t4["D146"],
                                             t4["C155"], t4["D155"]))]
    names = {"C": "V2 Off-cycle additions only",
             "D": f"V3 Extreme returns capped ({text_pct(WINSOR_LO)} and {text_pct(WINSOR_HI)})",
             "E": f"V4 Without {EXCL_YEAR_1} and {EXCL_YEAR_2}", "F": "V5 Sell at the close the day before inclusion",
             "G": f"V6 {TRIM_N} best and {TRIM_N} worst trades dropped"}
    for k, lab in names.items():
        a_, b_ = lambda o: t4[f"{k}{175 + o}"], lambda o: t4[f"{k}{198 + o}"]
        rows.append((lab, *c2(a_(0), b_(0), a_(2), a_(14), a_(5), a_(8), b_(2), b_(14), b_(5), b_(8), a_(17),
                              b_(17))))
    show("Table C2. Every variant tried, against the full sample",
         ["Trades", "A: median (p)", "A: average (p)", "B: median (p)", "B: average (p)", "Result"], rows,
         ["Each variant changes one rule. The average counts each announcement date once; medians use the Wilcoxon "
          "signed-rank test. " + star_note,
          f"V1's smallest average gain the test could detect: {pc(t4['C153'])} (A) and {pc(t4['D153'])} (B), with "
          f"{t4['C142']} announcement dates."])

    # ---------------------------------------------------------------------------------------------- Table D1
    k3 = ["C", "D", "E"]
    show("Table D1. Growth of each portfolio by half",
         ["Strategy, Window A", "Strategy, Window B", "Benchmark"],
         [(f"Value of $1 on {dmy(ex['C17'])}", *[usd(ex[k + "19"]) for k in k3]),
          (f"Growth a year, {dmy(first)} to {dmy(ex['C17'])} ({ex['C18']:,} days)",
           *[pc(ex[k + "20"], sign=True) for k in k3]),
          ("  above the Benchmark Portfolio", pc(ex["C24"], sign=True), pc(ex["D24"], sign=True), "-"),
          (f"Value of $1 on {dmy(last)}", *[usd(ex[k + "22"]) for k in k3]),
          (f"Growth a year, {dmy(to_date(da.date.iat[hday]))} to {dmy(last)} ({ex['C21']:,} days)",
           *[pc(ex[k + "23"], sign=True) for k in k3]),
          ("  above the Benchmark Portfolio", pc(ex["C25"], sign=True), pc(ex["D25"], sign=True), "-"),
          ("Second half: CAPM alpha a year (p)", f"{pc(ex['C26'], sign=True)} ({pv(ex['C27'])})",
           f"{pc(ex['D26'], sign=True)} ({pv(ex['D27'])})", "-")])

    # ---------------------------------------------------------------------------------------------- Table D2
    k4 = ["C", "D", "E", "F"]
    show("Table D2. A test for one break in the average trade, at a date chosen by the data",
         ["Window A", "Window B", f"A without {ex['E116']}", f"B without {ex['F116']}"],
         [("Announcement dates", *[str(ex[k + "104"]) for k in k4]),
          ("Largest F (sup-F)", *[nm(ex[k + "106"]) for k in k4]),
          ("Break after (last date before the break)", *[dmy(ex[k + "107"]) for k in k4]),
          ("Dates before / after the break", *[f"{ex[k + '109']} / {ex[k + '110']}" for k in k4]),
          ("Average before, each date counted once", *[pc(ex[k + "111"], sign=True) for k in k4]),
          ("Average after, each date counted once", *[pc(ex[k + "112"], sign=True) for k in k4]),
          ("Significant break?", *[ex[k + "115"] for k in k4])],
         [f"Critical values for the largest F (Bai and Perron, 1998, with {text_pct(BREAK_TRIM)} of dates kept out "
          f"at each end): {BAI_PERRON_CV_5} at 5% and {BAI_PERRON_CV_1} at 1%.",
          "Figure D1 plots the average trade on each announcement date (the trades are in trade_returns.csv, "
          "option --save) with the averages before and after the break."])


# =====================================================================================================================
# MAIN
# =====================================================================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(description="S&P 500 index-inclusion strategy: download the data, run every test "
                                             "and print the paper's tables.")
    ap.add_argument("--details", action="store_true", help="also print every statistic of every test, laid out like "
                                                           "the workbook's result sheets")
    ap.add_argument("--save", metavar="FOLDER", help="also save the workings as CSV files in FOLDER")
    args, _ = ap.parse_known_args(argv)              # ignores the extra arguments a notebook passes
    raw = download_data()
    print("\nPART 2. Running the calculations")
    res = run_calculations(raw)
    print_paper_tables(res)
    if args.details:
        print_workings(res.raw, res.tr, res.beta, res.hd, res.da)
        for rep in res.reports:
            print_report(rep)
    if args.save:
        save_workings(args.save, res.tr, res.beta, res.hd, res.da)
    return 0


if __name__ == "__main__":
    if "ipykernel" in sys.modules:                   # Colab or Jupyter (%run): no sys.exit, which would print a traceback
        main(sys.argv[1:])
    else:
        sys.exit(main())
