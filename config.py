import os

CHAT_ID_ALVO = "7078348831"

BASES_ACAREACOES = [
    ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_PASS")),
    ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_PASS")),
    ("GNH", os.getenv("IMILE_GNH_USER"), os.getenv("IMILE_PASS")),
    ("MNT", os.getenv("IMILE_MNT_USER"), os.getenv("IMILE_PASS")),
    ("GVR", os.getenv("IMILE_GVR_USER"), os.getenv("IMILE_PASS")),
    ("TFO", os.getenv("IMILE_TFO_USER"), os.getenv("IMILE_PASS")),
    ("RBN", os.getenv("IMILE_RBN_USER"), os.getenv("IMILE_PASS")),
    ("CPH", os.getenv("IMILE_CPH_USER"), os.getenv("IMILE_PASS")),
    ("QHG", os.getenv("IMILE_QHG_USER"), os.getenv("IMILE_PASS")),
    ("CTP", os.getenv("IMILE_CTP_USER"), os.getenv("IMILE_PASS"))
]

BASES_SLA = [
    ("JML", os.getenv("IMILE_JML_USER"), os.getenv("IMILE_PASS")),
    ("ITR", os.getenv("IMILE_ITR_USER"), os.getenv("IMILE_PASS"))
]

ITR_LOC_CITIES = ['ITABIRA']
ITR_INT_CITIES = ['BARAO DE COCAIS', 'SANTA BARBARA', 'CATAS ALTAS', 'ITAMBE DO MATO DENTRO', 'FERROS', 'MORRO DO PILAR', 'SANTO ANTONIO DO RIO ABAIXO', 'SANTA MARIA DE ITABIRA', 'PASSABEM', 'CARMESIA', 'SAO SEBASTIAO DO RIO PRETO']

JML_LOC_CITIES = ['BELA VISTA DE MINAS', 'DIONISIO', 'JOAO MONLEVADE', 'NOVA ERA', 'RIO PIRACICABA', 'SAO DOMINGOS DO PRATA', 'SAO GONCALO DO RIO ABAIXO', 'SAO JOSE DO GOIABAL']
JML_INT_CITIES = ['DOM SILVERIO', 'ALVINOPOLIS', 'BOM JESUS DO AMPARO', 'NOVA UNIAO', 'SEM-PEIXE', 'SEM PEIXE']

BACKLOG_CRITICO = 3

BACKLOG_FIM_DE_SEMANA = {"int": 4, "loc": 3}
BACKLOG_DIA_DE_SEMANA = {"int": 3, "loc": 2}
