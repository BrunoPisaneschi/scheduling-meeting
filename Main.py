# -*- coding:utf-8 -*-
#!/usr/bin/python3

import pathlib
from flow.FlowMaster import pipefy

if __name__ == '__main__':


    caminho_db = str(pathlib.Path(__file__).parent.absolute()) + "\\database\\"
    nome_db = "BASE CANDIDATOS"
    token = ""
    pipefy = pipefy(token, caminho_db, nome_db)
    pipefy.extract_datas()