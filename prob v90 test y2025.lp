\* china_coal_cost_problem *\
Minimize
OBJ: 140000000 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'HCC')
 + 60000000 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'PCI')
 + 60000000 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'SCC')
 + 40000000 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 + 48000000 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + 180310000 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'HCC')
 + 95430000 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'PCI')
 + 141890000 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'SCC')
 + 63160000 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 + 63160000 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 + 2994500 flow_total_('basi_idno_166',_'ctct_idno_5259')
 + 6896400 flow_total_('ctct_idno_4820',_'ctct_idno_5259')
 + 3825400 flow_total_('ctct_idno_4820',_'pwpl_idno_6749')
 + 2994500 flow_total_('ctct_idno_4820',_'rwst_idno_420')
 + 6896400 flow_total_('ctct_idno_5259',_'ctct_idno_4820')
 + 3234800 flow_total_('ctct_idno_5259',_'pwpl_idno_7198')
 + 3530300 flow_total_('ctct_idno_5259',_'stpt_idno_10708')
 + 5000000 flow_total_('ctct_idno_5259',_'stpt_idno_2')
 + 26950600 flow_total_('navo_idno_5467',_'navo_idno_5484')
 + 3886600 flow_total_('navo_idno_5467',_'port_idno_5584')
 + 26950600 flow_total_('navo_idno_5484',_'navo_idno_5467')
 + 3886600 flow_total_('port_idno_5584',_'navo_idno_5467')
 + 3743100 flow_total_('port_idno_5584',_'rwst_idno_508')
 + 2994500 flow_total_('rwst_idno_420',_'ctct_idno_4820')
 + 196100 flow_total_('rwst_idno_420',_'rwst_idno_508')
 + 3743100 flow_total_('rwst_idno_508',_'port_idno_5584')
 + 196100 flow_total_('rwst_idno_508',_'rwst_idno_420')
Subject To
_C1: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'HCC')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'HCC')
 + supply_coaltype_by_node_('basi_idno_166',_'HCC') >= 0
_C10: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_4500')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_4500')
 + supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_4500') >= 0
_C100: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 + supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_4500') >= 0
_C101: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 + supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_5500') >= 0
_C102: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 + supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_6500') >= 0
_C103: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'HCC')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 + supply_coaltype_by_node_('stpt_idno_10708',_'HCC') >= 0
_C104: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'PCI')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'PCI')
 + supply_coaltype_by_node_('stpt_idno_10708',_'PCI') >= 0
_C105: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'SCC')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'SCC')
 + supply_coaltype_by_node_('stpt_idno_10708',_'SCC') >= 0
_C106: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_4500')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 + supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_4500') >= 0
_C107: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_5500')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 + supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_5500') >= 0
_C108: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_6500')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 + supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_6500') >= 0
_C109: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'HCC')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 + supply_coaltype_by_node_('stpt_idno_2',_'HCC') >= 0
_C11: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_5500')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_5500')
 + supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_5500') >= 0
_C110: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'PCI')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'PCI')
 + supply_coaltype_by_node_('stpt_idno_2',_'PCI') >= 0
_C111: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'SCC')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'SCC')
 + supply_coaltype_by_node_('stpt_idno_2',_'SCC') >= 0
_C112: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_4500')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 + supply_coaltype_by_node_('stpt_idno_2',_'Thermal_4500') >= 0
_C113: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_5500')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 + supply_coaltype_by_node_('stpt_idno_2',_'Thermal_5500') >= 0
_C114: flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_6500')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 + supply_coaltype_by_node_('stpt_idno_2',_'Thermal_6500') >= 0
_C115: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'HCC')
 + supply_coaltype_by_node_('mine_idno_27',_'HCC') >= 0
_C116: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'PCI')
 + supply_coaltype_by_node_('mine_idno_27',_'PCI') >= 0
_C117: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'SCC')
 + supply_coaltype_by_node_('mine_idno_27',_'SCC') >= 0
_C118: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 + supply_coaltype_by_node_('mine_idno_27',_'Thermal_4500') >= 0
_C119: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 + supply_coaltype_by_node_('mine_idno_27',_'Thermal_5500') >= 0
_C12: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_6500')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_6500')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_6500')
 + supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_6500') >= 0
_C120: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_6500')
 + supply_coaltype_by_node_('mine_idno_27',_'Thermal_6500') >= 0
_C121: - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'HCC')
 + supply_coaltype_by_node_('mine_idno_11653',_'HCC') >= 0
_C122: - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'PCI')
 + supply_coaltype_by_node_('mine_idno_11653',_'PCI') >= 0
_C123: - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'SCC')
 + supply_coaltype_by_node_('mine_idno_11653',_'SCC') >= 0
_C124:
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_4500')
 + supply_coaltype_by_node_('mine_idno_11653',_'Thermal_4500') >= 0
_C125:
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 + supply_coaltype_by_node_('mine_idno_11653',_'Thermal_5500') >= 0
_C126:
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + supply_coaltype_by_node_('mine_idno_11653',_'Thermal_6500') >= 0
_C127:
 - 18.828 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('basi_idno_166',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('basi_idno_166',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('basi_idno_166',_'Thermal_6500') >= 0
_C128:
 - 18.828 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_6500') >= 2
_C129:
 18.828 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_6500') >= 2
_C13: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'HCC')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'HCC')
 + supply_coaltype_by_node_('ctct_idno_5259',_'HCC') >= 0
_C130:
 - 18.790344 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 - 22.965976 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 - 27.141608 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 + 18.790344 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 + 22.965976 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 + 27.141608 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 + 6.156756 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_4500')
 + 7.524924 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_5500')
 + 8.893092 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_6500')
 + 6.156756 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_4500')
 + 7.524924 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_5500')
 + 8.893092 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_6500') >= 2
_C131:
 18.790344 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 + 22.965976 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 + 27.141608 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 - 18.790344 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 - 22.965976 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 - 27.141608 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 + 5.686056 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_4500')
 + 6.949624 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_5500')
 + 8.213192 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_6500')
 + 5.686056 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_4500')
 + 6.949624 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_5500')
 + 8.213192 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_6500') >= 2
_C132:
 - 18.828 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('navo_idno_5467',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('navo_idno_5467',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('navo_idno_5467',_'Thermal_6500') >= 0
_C133:
 18.828 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('navo_idno_5484',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('navo_idno_5484',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('navo_idno_5484',_'Thermal_6500') >= 0
_C134:
 18.828 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('port_idno_5584',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('port_idno_5584',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('port_idno_5584',_'Thermal_6500') >= 0
_C135:
 18.828 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_6500') >= 0
_C136:
 18.828 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_6500') >= 0
_C137:
 18.828 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_6500')
 - 5.686056 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_4500')
 - 6.949624 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_5500')
 - 8.213192 flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_6500') >= 0
_C138:
 18.828 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_6500')
 - 5.686056 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_4500')
 - 6.949624 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_5500')
 - 8.213192 flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_6500') >= 0
_C139:
 18.828 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_6500')
 - 6.156756 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_4500')
 - 7.524924 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_5500')
 - 8.893092 flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_6500') >= 0
_C14: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'PCI')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'PCI')
 + supply_coaltype_by_node_('ctct_idno_5259',_'PCI') >= 0
_C140:
 18.828 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_6500')
 - 6.156756 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_4500')
 - 7.524924 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_5500')
 - 8.893092 flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_6500') >= 0
_C141:
 18.828 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('rwst_idno_420',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('rwst_idno_420',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('rwst_idno_420',_'Thermal_6500') >= 0
_C142:
 18.828 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('rwst_idno_508',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('rwst_idno_508',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('rwst_idno_508',_'Thermal_6500') >= 0
_C143:
 18.828 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 + 18.828 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_6500') >= 0
_C144:
 18.828 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_6500') >= 0
_C145:
 18.828 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_4500')
 + 23.012 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_5500')
 + 27.196 flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_6500')
 - 18.828 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('stpt_idno_2',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('stpt_idno_2',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('stpt_idno_2',_'Thermal_6500') >= 0
_C146:
 - 18.828 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('mine_idno_27',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('mine_idno_27',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('mine_idno_27',_'Thermal_6500') >= 0
_C147:
 - 18.828 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_4500')
 - 23.012 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 - 27.196 flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + 18.828 supply_coaltype_by_node_('mine_idno_11653',_'Thermal_4500')
 + 23.012 supply_coaltype_by_node_('mine_idno_11653',_'Thermal_5500')
 + 27.196 supply_coaltype_by_node_('mine_idno_11653',_'Thermal_6500') >= 0
_C148: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC') >= 13.929745
_C149:
 1.6694490818 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 - 5.49450549451 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'SCC')
 = 0
_C15: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'SCC')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'SCC')
 + supply_coaltype_by_node_('ctct_idno_5259',_'SCC') >= 0
_C150: 1.6694490818 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 - 5.49450549451 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'SCC')
 = 0
_C151:
 1.6694490818 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 - 5.40540540541 flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'PCI')
 = 0
_C152: 1.6694490818 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 - 5.40540540541 flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'PCI')
 = 0
_C153: flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'HCC')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'PCI')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'SCC')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500') <= 170
_C154: flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'HCC')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'PCI')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'SCC')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500') <= 170
_dummy: __dummy = 0
_C155: __dummy <= 4.641
_C156: __dummy <= 4.641
_C157: __dummy <= 5.628
_C158: __dummy <= 5.628
_C159:
 18.790344 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 + 22.965976 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 + 27.141608 flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 <= 99.3405
_C16: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_4500')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_4500')
 + supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_4500') >= 0
_C160:
 18.790344 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 + 22.965976 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 + 27.141608 flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 <= 99.3405
_C161: flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'HCC')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'PCI')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'SCC')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'HCC')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'PCI')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'SCC')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 + flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500') <= 1.4
_C162: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 + flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'PCI')
 + flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'SCC')
 + flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 + flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 + flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 <= 19.34
_C163: flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'PCI')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'SCC')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 <= 19.34
_C164: - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'HCC')
 - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'PCI')
 - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'SCC')
 - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 - flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_6500')
 + flow_total_('mine_idno_27',_'basi_idno_166') = 0
_C165: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'HCC')
 - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'PCI')
 - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'SCC')
 - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_4500')
 - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_5500')
 - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_6500')
 + flow_total_('basi_idno_166',_'ctct_idno_5259') = 0
_C166: - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_6500')
 + flow_total_('ctct_idno_5259',_'ctct_idno_4820') = 0
_C167: - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'HCC')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'PCI')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'SCC')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_6500')
 + flow_total_('ctct_idno_4820',_'rwst_idno_420') = 0
_C168: - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'HCC')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'PCI')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'SCC')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500')
 + flow_total_('rwst_idno_420',_'rwst_idno_508') = 0
_C169: - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'HCC')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'PCI')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'SCC')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_6500')
 + flow_total_('rwst_idno_508',_'port_idno_5584') = 0
_C17: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_5500')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_5500')
 + supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_5500') >= 0
_C170: - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'HCC')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'PCI')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'SCC')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 + flow_total_('port_idno_5584',_'navo_idno_5467') = 0
_C171: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'HCC')
 - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'PCI')
 - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'SCC')
 - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_4500')
 - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_5500')
 - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_6500')
 + flow_total_('navo_idno_5467',_'navo_idno_5484') = 0
_C172: - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'HCC')
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'PCI')
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'SCC')
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_4500')
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 - flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + flow_total_('mine_idno_11653',_'navo_idno_5484') = 0
_C173: - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_6500')
 + flow_total_('ctct_idno_5259',_'stpt_idno_10708') = 0
_C174: - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_6500')
 + flow_total_('ctct_idno_5259',_'stpt_idno_2') = 0
_C175: - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'PCI')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'SCC')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_4500')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_5500')
 - flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'Thermal_6500')
 + flow_total_('stpt_idno_10708',_'stdc_idno_12186') = 0
_C176: - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'PCI')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'SCC')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_4500')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_5500')
 - flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'Thermal_6500')
 + flow_total_('stpt_idno_2',_'stdc_idno_12186') = 0
_C177: - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'HCC')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'PCI')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'SCC')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_6500')
 + flow_total_('ctct_idno_4820',_'pwpl_idno_6749') = 0
_C178: - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'HCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'PCI')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'SCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_6500')
 + flow_total_('pwpl_idno_6749',_'pwun_idno_8687') = 0
_C179: - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'HCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'PCI')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'SCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_6500')
 + flow_total_('pwpl_idno_6749',_'pwun_idno_8688') = 0
_C18: flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_6500')
 + flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_10708',_'Thermal_6500')
 - flow_by_coaltype_('ctct_idno_5259',_'stpt_idno_2',_'Thermal_6500')
 + supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_6500') >= 0
_C180: - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'HCC')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'PCI')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'SCC')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_6500')
 + flow_total_('pwun_idno_8687',_'eldc_idno_11530') = 0
_C181: - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'HCC')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'PCI')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'SCC')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_6500')
 + flow_total_('pwun_idno_8688',_'eldc_idno_11530') = 0
_C182: - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'HCC')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'PCI')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'SCC')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_6500')
 + flow_total_('ctct_idno_5259',_'pwpl_idno_7198') = 0
_C183: - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'HCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'PCI')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'SCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_6500')
 + flow_total_('pwpl_idno_7198',_'pwun_idno_10114') = 0
_C184: - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'HCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'PCI')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'SCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_6500')
 + flow_total_('pwpl_idno_7198',_'pwun_idno_10115') = 0
_C185: - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'HCC')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'PCI')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'SCC')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_6500')
 + flow_total_('pwun_idno_10114',_'eldc_idno_11533') = 0
_C186: - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'HCC')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'PCI')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'SCC')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_6500')
 + flow_total_('pwun_idno_10115',_'eldc_idno_11533') = 0
_C187: - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'HCC')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'PCI')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'SCC')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 + flow_total_('eldc_idno_11533',_'eldc_idno_11530') = 0
_C188: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'HCC')
 - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'PCI')
 - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'SCC')
 - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_4500')
 - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_5500')
 - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'Thermal_6500')
 + flow_total_('ctct_idno_4820',_'ctct_idno_5259') = 0
_C189: - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'HCC')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'PCI')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'SCC')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_6500')
 + flow_total_('rwst_idno_420',_'ctct_idno_4820') = 0
_C19: - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'HCC')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'HCC')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'HCC')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'HCC')
 + supply_coaltype_by_node_('eldc_idno_11530',_'HCC') >= 0
_C190: - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'HCC')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'PCI')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'SCC')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500')
 + flow_total_('rwst_idno_508',_'rwst_idno_420') = 0
_C191: - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'HCC')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'PCI')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'SCC')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500')
 + flow_total_('port_idno_5584',_'rwst_idno_508') = 0
_C192: - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'HCC')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'PCI')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'SCC')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_4500')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_5500')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_6500')
 + flow_total_('navo_idno_5467',_'port_idno_5584') = 0
_C193: - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'HCC')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'PCI')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'SCC')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_4500')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_5500')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_6500')
 + flow_total_('navo_idno_5484',_'navo_idno_5467') = 0
_C194: - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'HCC')
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'PCI')
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'SCC')
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 + flow_total_('eldc_idno_11530',_'eldc_idno_11533') = 0
_C2: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'PCI')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'PCI')
 + supply_coaltype_by_node_('basi_idno_166',_'PCI') >= 0
_C20: - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'PCI')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'PCI')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'PCI')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'PCI')
 + supply_coaltype_by_node_('eldc_idno_11530',_'PCI') >= 0
_C21: - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'SCC')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'SCC')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'SCC')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'SCC')
 + supply_coaltype_by_node_('eldc_idno_11530',_'SCC') >= 0
_C22:
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_4500')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_4500')
 + supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_4500') >= 0
_C23:
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_5500')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_5500')
 + supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_5500') >= 0
_C24:
 - flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 + flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 + flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_6500')
 + flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_6500')
 + supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_6500') >= 0
_C25: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'HCC')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'HCC')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'HCC')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'HCC')
 + supply_coaltype_by_node_('eldc_idno_11533',_'HCC') >= 0
_C26: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'PCI')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'PCI')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'PCI')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'PCI')
 + supply_coaltype_by_node_('eldc_idno_11533',_'PCI') >= 0
_C27: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'SCC')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'SCC')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'SCC')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'SCC')
 + supply_coaltype_by_node_('eldc_idno_11533',_'SCC') >= 0
_C28: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_4500')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_4500')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_4500')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_4500')
 + supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_4500') >= 0
_C29: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_5500')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_5500')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_5500')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_5500')
 + supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_5500') >= 0
_C3: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'SCC')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'SCC')
 + supply_coaltype_by_node_('basi_idno_166',_'SCC') >= 0
_C30: flow_by_coaltype_('eldc_idno_11530',_'eldc_idno_11533',_'Thermal_6500')
 - flow_by_coaltype_('eldc_idno_11533',_'eldc_idno_11530',_'Thermal_6500')
 + flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_6500')
 + flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_6500')
 + supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_6500') >= 0
_C31: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'HCC')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'HCC')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'HCC')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'HCC')
 + supply_coaltype_by_node_('navo_idno_5467',_'HCC') >= 0
_C32: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'PCI')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'PCI')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'PCI')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'PCI')
 + supply_coaltype_by_node_('navo_idno_5467',_'PCI') >= 0
_C33: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'SCC')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'SCC')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'SCC')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'SCC')
 + supply_coaltype_by_node_('navo_idno_5467',_'SCC') >= 0
_C34: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_4500')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_4500')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_4500')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 + supply_coaltype_by_node_('navo_idno_5467',_'Thermal_4500') >= 0
_C35: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_5500')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_5500')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_5500')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 + supply_coaltype_by_node_('navo_idno_5467',_'Thermal_5500') >= 0
_C36: - flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_6500')
 - flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_6500')
 + flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_6500')
 + flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 + supply_coaltype_by_node_('navo_idno_5467',_'Thermal_6500') >= 0
_C37: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'HCC')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'HCC')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'HCC')
 + supply_coaltype_by_node_('navo_idno_5484',_'HCC') >= 0
_C38: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'PCI')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'PCI')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'PCI')
 + supply_coaltype_by_node_('navo_idno_5484',_'PCI') >= 0
_C39: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'SCC')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'SCC')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'SCC')
 + supply_coaltype_by_node_('navo_idno_5484',_'SCC') >= 0
_C4: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_4500')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_4500')
 + supply_coaltype_by_node_('basi_idno_166',_'Thermal_4500') >= 0
_C40: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_4500')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_4500')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_4500')
 + supply_coaltype_by_node_('navo_idno_5484',_'Thermal_4500') >= 0
_C41: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_5500')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_5500')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_5500')
 + supply_coaltype_by_node_('navo_idno_5484',_'Thermal_5500') >= 0
_C42: flow_by_coaltype_('mine_idno_11653',_'navo_idno_5484',_'Thermal_6500')
 + flow_by_coaltype_('navo_idno_5467',_'navo_idno_5484',_'Thermal_6500')
 - flow_by_coaltype_('navo_idno_5484',_'navo_idno_5467',_'Thermal_6500')
 + supply_coaltype_by_node_('navo_idno_5484',_'Thermal_6500') >= 0
_C43: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'HCC')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'HCC')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'HCC')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'HCC')
 + supply_coaltype_by_node_('port_idno_5584',_'HCC') >= 0
_C44: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'PCI')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'PCI')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'PCI')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'PCI')
 + supply_coaltype_by_node_('port_idno_5584',_'PCI') >= 0
_C45: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'SCC')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'SCC')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'SCC')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'SCC')
 + supply_coaltype_by_node_('port_idno_5584',_'SCC') >= 0
_C46: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_4500')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_4500')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_4500')
 + supply_coaltype_by_node_('port_idno_5584',_'Thermal_4500') >= 0
_C47: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_5500')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_5500')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_5500')
 + supply_coaltype_by_node_('port_idno_5584',_'Thermal_5500') >= 0
_C48: flow_by_coaltype_('navo_idno_5467',_'port_idno_5584',_'Thermal_6500')
 - flow_by_coaltype_('port_idno_5584',_'navo_idno_5467',_'Thermal_6500')
 - flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500')
 + flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_6500')
 + supply_coaltype_by_node_('port_idno_5584',_'Thermal_6500') >= 0
_C49: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'HCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'HCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'HCC')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'HCC') >= 0
_C5: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_5500')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_5500')
 + supply_coaltype_by_node_('basi_idno_166',_'Thermal_5500') >= 0
_C50: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'PCI')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'PCI')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'PCI')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'PCI') >= 0
_C51: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'SCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'SCC')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'SCC')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'SCC') >= 0
_C52: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_4500')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_4500') >= 0
_C53: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_5500')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_5500') >= 0
_C54: flow_by_coaltype_('ctct_idno_5259',_'pwpl_idno_7198',_'Thermal_6500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_6500')
 - flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_6500')
 + supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_6500') >= 0
_C55: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'HCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'HCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'HCC')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'HCC') >= 0
_C56: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'PCI')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'PCI')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'PCI')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'PCI') >= 0
_C57: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'SCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'SCC')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'SCC')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'SCC') >= 0
_C58: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_4500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_4500')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_4500') >= 0
_C59: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_5500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_5500')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_5500') >= 0
_C6: - flow_by_coaltype_('basi_idno_166',_'ctct_idno_5259',_'Thermal_6500')
 + flow_by_coaltype_('mine_idno_27',_'basi_idno_166',_'Thermal_6500')
 + supply_coaltype_by_node_('basi_idno_166',_'Thermal_6500') >= 0
_C60: flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'Thermal_6500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_6500')
 - flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_6500')
 + supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_6500') >= 0
_C61: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'HCC')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'HCC')
 + supply_coaltype_by_node_('pwun_idno_10114',_'HCC') >= 0
_C62: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'PCI')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'PCI')
 + supply_coaltype_by_node_('pwun_idno_10114',_'PCI') >= 0
_C63: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'SCC')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'SCC')
 + supply_coaltype_by_node_('pwun_idno_10114',_'SCC') >= 0
_C64: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_4500')
 + supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_4500') >= 0
_C65: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_5500')
 + supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_5500') >= 0
_C66: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10114',_'Thermal_6500')
 - flow_by_coaltype_('pwun_idno_10114',_'eldc_idno_11533',_'Thermal_6500')
 + supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_6500') >= 0
_C67: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'HCC')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'HCC')
 + supply_coaltype_by_node_('pwun_idno_10115',_'HCC') >= 0
_C68: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'PCI')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'PCI')
 + supply_coaltype_by_node_('pwun_idno_10115',_'PCI') >= 0
_C69: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'SCC')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'SCC')
 + supply_coaltype_by_node_('pwun_idno_10115',_'SCC') >= 0
_C7: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'HCC')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'HCC')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'HCC')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'HCC')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'HCC')
 + supply_coaltype_by_node_('ctct_idno_4820',_'HCC') >= 0
_C70: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_4500')
 + supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_4500') >= 0
_C71: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_5500')
 + supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_5500') >= 0
_C72: flow_by_coaltype_('pwpl_idno_7198',_'pwun_idno_10115',_'Thermal_6500')
 - flow_by_coaltype_('pwun_idno_10115',_'eldc_idno_11533',_'Thermal_6500')
 + supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_6500') >= 0
_C73: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'HCC')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'HCC')
 + supply_coaltype_by_node_('pwun_idno_8687',_'HCC') >= 0
_C74: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'PCI')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'PCI')
 + supply_coaltype_by_node_('pwun_idno_8687',_'PCI') >= 0
_C75: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'SCC')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'SCC')
 + supply_coaltype_by_node_('pwun_idno_8687',_'SCC') >= 0
_C76: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_4500')
 + supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_4500') >= 0
_C77: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_5500')
 + supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_5500') >= 0
_C78: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8687',_'Thermal_6500')
 - flow_by_coaltype_('pwun_idno_8687',_'eldc_idno_11530',_'Thermal_6500')
 + supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_6500') >= 0
_C79: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'HCC')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'HCC')
 + supply_coaltype_by_node_('pwun_idno_8688',_'HCC') >= 0
_C8: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'PCI')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'PCI')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'PCI')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'PCI')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'PCI')
 + supply_coaltype_by_node_('ctct_idno_4820',_'PCI') >= 0
_C80: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'PCI')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'PCI')
 + supply_coaltype_by_node_('pwun_idno_8688',_'PCI') >= 0
_C81: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'SCC')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'SCC')
 + supply_coaltype_by_node_('pwun_idno_8688',_'SCC') >= 0
_C82: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_4500')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_4500')
 + supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_4500') >= 0
_C83: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_5500')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_5500')
 + supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_5500') >= 0
_C84: flow_by_coaltype_('pwpl_idno_6749',_'pwun_idno_8688',_'Thermal_6500')
 - flow_by_coaltype_('pwun_idno_8688',_'eldc_idno_11530',_'Thermal_6500')
 + supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_6500') >= 0
_C85: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'HCC')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'HCC')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'HCC')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'HCC')
 + supply_coaltype_by_node_('rwst_idno_420',_'HCC') >= 0
_C86: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'PCI')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'PCI')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'PCI')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'PCI')
 + supply_coaltype_by_node_('rwst_idno_420',_'PCI') >= 0
_C87: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'SCC')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'SCC')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'SCC')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'SCC')
 + supply_coaltype_by_node_('rwst_idno_420',_'SCC') >= 0
_C88: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 + supply_coaltype_by_node_('rwst_idno_420',_'Thermal_4500') >= 0
_C89: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 + supply_coaltype_by_node_('rwst_idno_420',_'Thermal_5500') >= 0
_C9: - flow_by_coaltype_('ctct_idno_4820',_'ctct_idno_5259',_'SCC')
 - flow_by_coaltype_('ctct_idno_4820',_'pwpl_idno_6749',_'SCC')
 - flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'SCC')
 + flow_by_coaltype_('ctct_idno_5259',_'ctct_idno_4820',_'SCC')
 + flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'SCC')
 + supply_coaltype_by_node_('ctct_idno_4820',_'SCC') >= 0
_C90: flow_by_coaltype_('ctct_idno_4820',_'rwst_idno_420',_'Thermal_6500')
 - flow_by_coaltype_('rwst_idno_420',_'ctct_idno_4820',_'Thermal_6500')
 - flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500')
 + flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500')
 + supply_coaltype_by_node_('rwst_idno_420',_'Thermal_6500') >= 0
_C91: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'HCC')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'HCC')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'HCC')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'HCC')
 + supply_coaltype_by_node_('rwst_idno_508',_'HCC') >= 0
_C92: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'PCI')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'PCI')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'PCI')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'PCI')
 + supply_coaltype_by_node_('rwst_idno_508',_'PCI') >= 0
_C93: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'SCC')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'SCC')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'SCC')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'SCC')
 + supply_coaltype_by_node_('rwst_idno_508',_'SCC') >= 0
_C94: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_4500')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_4500')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_4500')
 + supply_coaltype_by_node_('rwst_idno_508',_'Thermal_4500') >= 0
_C95: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_5500')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_5500')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_5500')
 + supply_coaltype_by_node_('rwst_idno_508',_'Thermal_5500') >= 0
_C96: flow_by_coaltype_('port_idno_5584',_'rwst_idno_508',_'Thermal_6500')
 + flow_by_coaltype_('rwst_idno_420',_'rwst_idno_508',_'Thermal_6500')
 - flow_by_coaltype_('rwst_idno_508',_'port_idno_5584',_'Thermal_6500')
 - flow_by_coaltype_('rwst_idno_508',_'rwst_idno_420',_'Thermal_6500')
 + supply_coaltype_by_node_('rwst_idno_508',_'Thermal_6500') >= 0
_C97: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'HCC')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'HCC')
 + supply_coaltype_by_node_('stdc_idno_12186',_'HCC') >= 0
_C98: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'PCI')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'PCI')
 + supply_coaltype_by_node_('stdc_idno_12186',_'PCI') >= 0
_C99: flow_by_coaltype_('stpt_idno_10708',_'stdc_idno_12186',_'SCC')
 + flow_by_coaltype_('stpt_idno_2',_'stdc_idno_12186',_'SCC')
 + supply_coaltype_by_node_('stdc_idno_12186',_'SCC') >= 0
Bounds
__dummy = 0
supply_coaltype_by_node_('basi_idno_166',_'HCC') = 0
supply_coaltype_by_node_('basi_idno_166',_'PCI') = 0
supply_coaltype_by_node_('basi_idno_166',_'SCC') = 0
supply_coaltype_by_node_('basi_idno_166',_'Thermal_4500') = 0
supply_coaltype_by_node_('basi_idno_166',_'Thermal_5500') = 0
supply_coaltype_by_node_('basi_idno_166',_'Thermal_6500') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'HCC') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'PCI') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'SCC') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_4500') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_5500') = 0
supply_coaltype_by_node_('ctct_idno_4820',_'Thermal_6500') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'HCC') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'PCI') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'SCC') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_4500') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_5500') = 0
supply_coaltype_by_node_('ctct_idno_5259',_'Thermal_6500') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'HCC') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'PCI') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'SCC') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_4500') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_5500') = 0
supply_coaltype_by_node_('eldc_idno_11530',_'Thermal_6500') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'HCC') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'PCI') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'SCC') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_4500') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_5500') = 0
supply_coaltype_by_node_('eldc_idno_11533',_'Thermal_6500') = 0
supply_coaltype_by_node_('mine_idno_11653',_'HCC') <= 1000000
supply_coaltype_by_node_('mine_idno_11653',_'PCI') <= 1000000
supply_coaltype_by_node_('mine_idno_11653',_'SCC') <= 1000000
supply_coaltype_by_node_('mine_idno_11653',_'Thermal_4500') = 0
supply_coaltype_by_node_('mine_idno_11653',_'Thermal_5500') <= 10
supply_coaltype_by_node_('mine_idno_11653',_'Thermal_6500') <= 5
supply_coaltype_by_node_('mine_idno_27',_'HCC') <= 1000000
supply_coaltype_by_node_('mine_idno_27',_'PCI') <= 1000000
supply_coaltype_by_node_('mine_idno_27',_'SCC') <= 1000000
supply_coaltype_by_node_('mine_idno_27',_'Thermal_4500') <= 5
supply_coaltype_by_node_('mine_idno_27',_'Thermal_5500') <= 10
supply_coaltype_by_node_('mine_idno_27',_'Thermal_6500') = 0
supply_coaltype_by_node_('navo_idno_5467',_'HCC') = 0
supply_coaltype_by_node_('navo_idno_5467',_'PCI') = 0
supply_coaltype_by_node_('navo_idno_5467',_'SCC') = 0
supply_coaltype_by_node_('navo_idno_5467',_'Thermal_4500') = 0
supply_coaltype_by_node_('navo_idno_5467',_'Thermal_5500') = 0
supply_coaltype_by_node_('navo_idno_5467',_'Thermal_6500') = 0
supply_coaltype_by_node_('navo_idno_5484',_'HCC') = 0
supply_coaltype_by_node_('navo_idno_5484',_'PCI') = 0
supply_coaltype_by_node_('navo_idno_5484',_'SCC') = 0
supply_coaltype_by_node_('navo_idno_5484',_'Thermal_4500') = 0
supply_coaltype_by_node_('navo_idno_5484',_'Thermal_5500') = 0
supply_coaltype_by_node_('navo_idno_5484',_'Thermal_6500') = 0
supply_coaltype_by_node_('port_idno_5584',_'HCC') = 0
supply_coaltype_by_node_('port_idno_5584',_'PCI') = 0
supply_coaltype_by_node_('port_idno_5584',_'SCC') = 0
supply_coaltype_by_node_('port_idno_5584',_'Thermal_4500') = 0
supply_coaltype_by_node_('port_idno_5584',_'Thermal_5500') = 0
supply_coaltype_by_node_('port_idno_5584',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'HCC') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'PCI') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'SCC') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwpl_idno_6749',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'HCC') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'PCI') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'SCC') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwpl_idno_7198',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'HCC') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'PCI') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'SCC') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwun_idno_10114',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'HCC') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'PCI') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'SCC') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwun_idno_10115',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'HCC') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'PCI') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'SCC') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwun_idno_8687',_'Thermal_6500') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'HCC') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'PCI') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'SCC') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_4500') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_5500') = 0
supply_coaltype_by_node_('pwun_idno_8688',_'Thermal_6500') = 0
supply_coaltype_by_node_('rwst_idno_420',_'HCC') = 0
supply_coaltype_by_node_('rwst_idno_420',_'PCI') = 0
supply_coaltype_by_node_('rwst_idno_420',_'SCC') = 0
supply_coaltype_by_node_('rwst_idno_420',_'Thermal_4500') = 0
supply_coaltype_by_node_('rwst_idno_420',_'Thermal_5500') = 0
supply_coaltype_by_node_('rwst_idno_420',_'Thermal_6500') = 0
supply_coaltype_by_node_('rwst_idno_508',_'HCC') = 0
supply_coaltype_by_node_('rwst_idno_508',_'PCI') = 0
supply_coaltype_by_node_('rwst_idno_508',_'SCC') = 0
supply_coaltype_by_node_('rwst_idno_508',_'Thermal_4500') = 0
supply_coaltype_by_node_('rwst_idno_508',_'Thermal_5500') = 0
supply_coaltype_by_node_('rwst_idno_508',_'Thermal_6500') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'HCC') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'PCI') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'SCC') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_4500') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_5500') = 0
supply_coaltype_by_node_('stdc_idno_12186',_'Thermal_6500') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'HCC') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'PCI') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'SCC') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_4500') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_5500') = 0
supply_coaltype_by_node_('stpt_idno_10708',_'Thermal_6500') = 0
supply_coaltype_by_node_('stpt_idno_2',_'HCC') = 0
supply_coaltype_by_node_('stpt_idno_2',_'PCI') = 0
supply_coaltype_by_node_('stpt_idno_2',_'SCC') = 0
supply_coaltype_by_node_('stpt_idno_2',_'Thermal_4500') = 0
supply_coaltype_by_node_('stpt_idno_2',_'Thermal_5500') = 0
supply_coaltype_by_node_('stpt_idno_2',_'Thermal_6500') = 0
End
