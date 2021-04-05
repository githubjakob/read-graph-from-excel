{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 11,
   "id": "quarterly-friend",
   "metadata": {},
   "outputs": [],
   "source": [
    "from openpyxl import Workbook\n",
    "from openpyxl import load_workbook"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 56,
   "id": "legal-variance",
   "metadata": {},
   "outputs": [],
   "source": [
    "from openpyxl.formula import Tokenizer"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 34,
   "id": "pressing-reservoir",
   "metadata": {},
   "outputs": [],
   "source": [
    "from graphviz import Digraph"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 79,
   "id": "indian-julian",
   "metadata": {},
   "outputs": [],
   "source": [
    "dot = Digraph(comment='Excel')"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 80,
   "id": "continuing-second",
   "metadata": {},
   "outputs": [],
   "source": [
    "wb = load_workbook(filename = '/home/jakob/projects/read-graph-from-excel/example.xlsx')"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 43,
   "id": "occasional-ranking",
   "metadata": {},
   "outputs": [],
   "source": [
    "LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'\n",
    "def excel_style(row, col):\n",
    "    \"\"\" Convert given row and column number to an Excel-style cell name. \"\"\"\n",
    "    result = []\n",
    "    while col:\n",
    "        col, rem = divmod(col-1, 26)\n",
    "        result[:0] = LETTERS[rem]\n",
    "    return ''.join(result) + str(row)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 81,
   "id": "affected-party",
   "metadata": {},
   "outputs": [],
   "source": [
    "max_col = 3 ## A = 1, 1 based\n",
    "max_row = 5 ## 1 based"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 82,
   "id": "continent-isolation",
   "metadata": {},
   "outputs": [],
   "source": [
    "sheet_ranges = wb['Sheet1']"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 83,
   "id": "limited-compatibility",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "B1==A1+1\n",
      "B1 has a reference to A1\n",
      "C1==B1+1\n",
      "C1 has a reference to B1\n",
      "B2==C1*2\n",
      "B2 has a reference to C1\n",
      "B5==IF(A5>500,B2+2,\"Low\")\n",
      "B5 has a reference to A5\n",
      "B5 has a reference to B2\n"
     ]
    }
   ],
   "source": [
    "for y in range(1, max_row+1):\n",
    "    for x in range(2,max_col+1):\n",
    "        value = sheet_ranges.cell(y,x).value\n",
    "        cell = excel_style(y,x)\n",
    "        if value is not None:\n",
    "            print(f\"{cell}={value}\")\n",
    "            tok = Tokenizer(value)\n",
    "            ##print(\"\\n\".join(\"%12s%11s%9s\" % (t.value, t.type, t.subtype) for t in tok.items))\n",
    "            for t in tok.items:\n",
    "                if t.subtype == \"RANGE\":\n",
    "                    print(f\"{cell} has a reference to {t.value}\")\n",
    "                    dot.edge(cell, t.value, constraint='false')\n",
    "            dot.node(cell, f\"{cell}={value}\")"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 84,
   "id": "sharp-crack",
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "'/home/jakob/projects/read-graph-from-excel/graph.gv.pdf'"
      ]
     },
     "execution_count": 84,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "dot.render(\"/home/jakob/projects/read-graph-from-excel/graph.gv\", view=True)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 77,
   "id": "legendary-acquisition",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "// Excel\n",
      "digraph {\n",
      "\tB1 -> A1 [constraint=false]\n",
      "\tB1 [label=\"B1==A1+1\"]\n",
      "\tC1 -> B1 [constraint=false]\n",
      "\tC1 [label=\"C1==B1+1\"]\n",
      "\tB2 -> C1 [constraint=false]\n",
      "\tB2 [label=\"B2==C1*2\"]\n",
      "}\n"
     ]
    }
   ],
   "source": [
    "print(dot.source)  "
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.8.5"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
