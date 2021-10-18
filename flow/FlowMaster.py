# -*- coding:utf-8 -*-
# /usr/bin/python3

import os
import re
from datetime import datetime, timezone
import requests
from unidecode import unidecode
from utils.Pipefy import Pipefy
from utils.ExcelUtils import excel
from utils.Google_Calendar_API import main
import base64
from os import remove


class pipefy():
    def __init__(self, token, caminho_db, nome_db):
        self._token = token
        self._pipe_id = 1102385
        self._dict_datas = {}
        self._db_utils = excel(caminho_db, nome_db)

    def extract_datas(self):
        _pipefy = Pipefy(self._token)
        _pipes = _pipefy.pipes([self._pipe_id])[0]
        _count = 1
        for _cards in _pipes['phases']:
            if 'entrevista rh' in unidecode(_cards['name'].lower()) or 'entrevista ta(c)cnica' in unidecode(
                    _cards['name'].lower()) or 'entrevista cliente' in unidecode(_cards['name'].lower()):
            # if unidecode(_cards['name'].lower()) == 'entrevista cliente':
                for _edge in _cards['cards']['edges']:
                    self._dict_datas[_count] = {'id_card': '',
                                                'titulo': '',
                                                'vaga': '',
                                                'nivel': '',
                                                'motivo_de_recusa': '',
                                                'como_soube_da_vaga': '',
                                                'email_candidato': '',
                                                'status_card': '',
                                                'email_responsavel': '',
                                                'data_entrevista': '',
                                                'horario_entrevista': '',
                                                'link_entrevista': '',
                                                'curriculo_candidato': '',
                                                'canal_contato': '',
                                                'email_cliente': '',
                                                'data_criacao': '',
                                                'data_modificacao': '',
                                                'linkedin': '',
                                                'status': '1'
                                                }
                    print("Id do card: {}".format(_edge['node']['id']))
                    print("Titulo do card: {}".format(_edge['node']['title']))
                    self._dict_datas[_count].update({'id_card': _edge['node']['id'], 'titulo': _edge['node']['title']})
                    _infos_card = _pipefy.card(_edge['node']['id'])
                    self._dict_datas[_count].update({'status_card': _infos_card['current_phase']['name'],
                                                     'data_criacao': self._convert_time(
                                                         _infos_card['phases_history'][0]['firstTimeIn']),
                                                     'data_modificacao': self._convert_time(
                                                         _infos_card['phases_history'][-1]['firstTimeIn'])})

                    for _fields in _infos_card['fields']:
                        if (_fields['name'].lower() == 'e-mail'):
                            print("Email: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'email_candidato': _fields['value'].strip()})

                        elif ('linkedin' in _fields['name'].lower()):
                            print("Email: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'linkedin': _fields['value'].strip()})

                        elif ('currículo' in _fields['name'].lower() and _fields['value'].strip() != '[]'):
                            print("Currículo: {}".format(re.sub(r'[\[\]\\"]', '', _fields['value'].strip())))
                            self._dict_datas[_count].update({'curriculo_candidato': self._get_base64_curriculo(
                                re.sub(r'[\[\]\\"]', '', _fields['value'].strip()))})

                        elif ('horário entrevista' in _fields['name'].lower()):
                            _extract_date = _fields['value'].strip().split(" ")
                            _interview_date = _extract_date[0]
                            _interview_time = _extract_date[1]
                            print("Data Entrevista: {}".format(_interview_date))
                            print("Horário Entrevista: {}".format(_interview_time))
                            self._dict_datas[_count].update(
                                {'data_entrevista': _interview_date, 'horario_entrevista': _interview_time})

                        elif ('e-mail do responsável' in _fields['name'].lower()):
                            print("Email do responsável: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'email_responsavel': _fields['value'].strip()})

                        elif ('responsável pela entrevista - rh' in _fields['name'].lower()):
                            print("Email: {}".format(_fields['value'].strip()))
                            self._dict_datas[_count].update({'email_cliente': _fields['value'].strip()})
                    print("-" * 20)
                    if (self._dict_datas[_count]['status_card'] == 'entrevista cliente' and not
                    self._dict_datas[_count]['email_cliente']):
                        del self._dict_datas[_count]
                        continue
                    if (not self._dict_datas[_count]['data_entrevista'] or
                            not self._dict_datas[_count]['email_candidato'] or
                            not self._dict_datas[_count]['email_responsavel']):
                        del self._dict_datas[_count]
                        continue
                    _exist_card = self._consult_db(self._dict_datas[_count])
                    if (_exist_card):
                        del self._dict_datas[_count]
                        continue

                    self._db_utils.update_id(self._dict_datas[_count])

                    _count += 1

    def _consult_db(self, _data):
        _dict_db = self._db_utils.read_excel(db=True)
        for _index_db, _row_db in _dict_db.items():
            if (int(_row_db['id_card']) == int(_data['id_card'])) and \
                    (_row_db['status_card'] == _data['status_card']) and \
                    (_row_db['data_modificacao'] == _data['data_modificacao']):
                return True
        return False

    def _utc_to_local(self, utc_dt):
        return utc_dt.replace(tzinfo=timezone.utc).astimezone(tz=None)

    def _convert_time(self, _string_time):
        return self._utc_to_local(datetime.fromisoformat(_string_time)).strftime("%Y-%m-%d %H:%M:%S")

    def _get_base64_curriculo(self, _string_link):
        _headers = {'Content-Type': 'text/plain'}
        response = requests.get(_string_link)
        _caminho_curriculo = r'.\files\curriculo.pdf'
        if (not os.path.exists(r'.\files')):
            os.makedirs(r'.\files')
        _file = open(_caminho_curriculo, 'wb')
        _file.write(response.content)
        _file.close()
        _data = open(_caminho_curriculo, "rb").read()
        _base64 = base64.b64encode(_data).decode("utf-8")
        remove(_caminho_curriculo)
        return _base64

    def _create_invite(self):
        main()
