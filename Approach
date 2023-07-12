#!IMPORTANT Validation and simulation
import networkx as nx
import xlrd
import xlsxwriter
import numpy as np

#Column A is a list of components whhich have a direct link to components in column B.
#Each raw indicates a direct link between component A and component B.
#Column C indicates the correlation Coefficinets (link information) of the link.
#Alternatively you may use Core.xlsx, CoreFamilyEducation.xlsx and CoreFamily.xlsx
file = "Whole.xlsx"

_book = xlrd.open_workbook(file)
_sheet = _book.sheet_by_index(1)

allOutgoingNodes = list()
noRepeatitionNodes = list()

for row in range(_sheet.nrows):
    if row > 0:
        _data = _sheet.row_slice(row)
        allOutgoingNodes.append(_data[0].value)

noRepeatitionNodes = list(dict.fromkeys(allOutgoingNodes))

print(noRepeatitionNodes)
flag = 1
pr = None
for _nodes in noRepeatitionNodes:
    j = 1
    occurrences = allOutgoingNodes.count(_nodes)
    workbook = xlsxwriter.Workbook("simulation_Output/" + _nodes + '_' + file)
    worksheet = workbook.add_worksheet()

    for w in np.arange(1, 100, 1):
        print(w)
        G = nx.DiGraph()
        weight = 0
        for row in range(_sheet.nrows):
            if row > 0:
                _data = _sheet.row_slice(row)
                _Node1 = _data[0].value
                _Node2 = _data[1].value
                if _Node2 == _nodes:
                    weight = float(_data[2].value) * float(w)
                elif _Node2 != _nodes:
                    weight = float(_data[2].value)
                G.add_weighted_edges_from([(_Node1, _Node2, weight)])

        nodes_list = list(G.nodes())
        pr = nx.pagerank_numpy(G, alpha=0.65)
        pr_list = list(pr)
        pr_keys = list(pr.keys())
        pr_values = list(pr.values())
        my_formatted_list = ['%.4f' % elem for elem in pr_values]
        if w == 1:
            for col_num, data in enumerate(pr_keys):
                worksheet.write(0, col_num, data)
        for col_num, data in enumerate(my_formatted_list):
            worksheet.write(j, col_num, float(data))
        j += 1
        G.clear()
        # pr.clear()
        del G
        # del pr
    workbook.close()
