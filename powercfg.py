import subprocess
import chardet
import re
from abc import ABC, abstractmethod
import json
from typing import Any
import os.path

class WrongSettingValueException(Exception): ...

class Node(ABC):
    _GUID_PATTERN = r'\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b'
    _NAME_PATTERN = r'\(([^)]+)\)'
    _SPACE_2_PATTERN = r'^ {2}[^ ]'
    _SPACE_4_PATTERN = r'^ {4}[^ ]'
    _SPACE_6_PATTERN = r'^ {6}[^ ]'

    _guid : str = None
    _name : str = None

    def __init__(self, str_to_parse:str) -> None:
        super().__init__()
        self.__start_parse(str_to_parse)

    @abstractmethod
    def _parse(self, rows:list[str]):
        pass

    @abstractmethod
    def load_from_json(self, json_data:dict):
        pass

    @abstractmethod
    def to_json(self) -> Any:
        pass

    def __start_parse(self, str_to_parse:str):
        rows = str_to_parse.split('\n')

        # find GUID and name in the first row
        self._guid = self._find_str(rows[0], self._GUID_PATTERN)
        self._name = self._find_str(rows[0], self._NAME_PATTERN)
        if self._name != None:
            self._name = self._name[1:-1]

        self._parse(rows[1:])

    def _find_str(self, string:str, pattern:str) -> str:
        match = re.search(pattern, string)
        if match:
            return match.group(0)
        return None
    
    def _find_index(self, string:str, pattern:str) -> int:
        match = re.search(pattern, string)
        if match:
            return match.span(0)
        return -1

    def get_guid(self) -> str:
        return self._guid
    
    def get_name(self) -> str:
        return self._name

class Setting(Node):
    RANGE_OPTIONS : int = 0
    LIST_OPTIONS : int = 1

    __doc : list[dict] = None

    __ac_value : int = None
    __old_ac_value : int = None

    __dc_value : int = None
    __old_dc_value : int = None

    __options_type : int = None
    __options : list[int] = None
    
    def __init__(self, str_to_parse: str) -> None:
        super().__init__(str_to_parse)
        self.__parse_options()

    def __value_parse(self, value:str) -> str:
        if value.startswith('0x'):
            return str(int(value, 16))
        else:
            return value
        
    def __parse_options(self) -> None:
        val_1:str = self.__doc[0]['value']
        val_2:str = self.__doc[1]['value']

        if val_1.isdigit() and val_2.isdigit():
            self.__options_type = Setting.RANGE_OPTIONS
            self.__options = [int(val_1), int(val_2)]
        else:
            self.__options_type = Setting.LIST_OPTIONS
            self.__options = []
            for item in self.__doc:
                if item['value'].isdigit():
                    self.__options.append(int(item['value']))

    def __check_value(self, value:int) -> bool:
        if self.__options_type == Setting.RANGE_OPTIONS:
            return value >= self.__options[0] and value <= self.__options[1]
        else:
            return value in self.__options
    
    def _parse(self, rows:list[str]):
        self.__doc = []
        for i in range(len(rows)):
            # if row contains doc
            if self._find_index(rows[i], self._SPACE_6_PATTERN) != -1 and rows[i].find('GUID') == -1:
                spt = rows[i].strip().split(':')
                self.__doc.append({
                    'description': spt[0],
                    'value': self.__value_parse(spt[1].strip())
                })
            # if row contains AC/DC value
            elif self._find_index(rows[i], self._SPACE_4_PATTERN) != -1:
                value = int(rows[i].split(':')[1].strip(), 16)
                if self.__ac_value == None:
                    self.__ac_value = value
                    self.__old_ac_value = value
                else:
                    self.__dc_value = value
                    self.__old_dc_value = value

    def load_from_json(self, json_data: dict):
        self.set_ac_value(json_data['ac_value'])
        self.set_dc_value(json_data['dc_value'])

    def get_ac_value(self) -> int:
        return self.__ac_value

    def get_ac_value_hex(self) -> str:
        return hex(self.__ac_value)
    
    def get_dc_value(self) -> int:
        return self.__dc_value
    
    def get_dc_value_hex(self) -> str:
        return hex(self.__dc_value)
    
    def get_doc(self) -> list[dict]:
        return self.__doc
    
    def get_options_type(self) -> int:
        return self.__options_type
    
    def get_options_type_str(self) -> str:
        if self.get_options_type() == Setting.RANGE_OPTIONS:
            return 'RANGE'
        else:
            return 'LIST'
    
    def get_options(self) -> list[int]:
        return self.__options
    
    def __set_value(self, value:int, is_ac:bool) -> None:
        if self.__check_value(value):
            if is_ac:
                self.__ac_value = value
            else:
                self.__dc_value = value
        else:
            m = f'Setting: {self.get_name()}; GUID: {self.get_guid()}; value to set: {value}; available options: {self.get_options()}; options type: {self.get_options_type_str()}'
            raise WrongSettingValueException(m)
    
    def set_ac_value(self, ac_value:int) -> None:
        self.__set_value(ac_value, True)
        
    def set_dc_value(self, dc_value:int) -> None:
        self.__set_value(dc_value, False)

    def is_ac_changed(self) -> bool:
        return self.__ac_value != self.__old_ac_value
    
    def is_dc_changed(self) -> bool:
        return self.__dc_value != self.__old_dc_value
    
    def update_old_values(self):
        self.__old_ac_value = self.__ac_value
        self.__old_dc_value = self.__dc_value

    def to_json(self) -> Any:
        return {
            'name': self.get_name(),
            'options_type': self.get_options_type(),
            'options': self.get_options(),
            'ac_value': self.get_ac_value(),
            'dc_value': self.get_dc_value(),
            'doc': self.get_doc()
        }

class SubGroup(Node):
    __settings : list[Setting] = None

    def __init__(self, str_to_parse: str) -> None:
        super().__init__(str_to_parse)

    def _parse(self, rows:list[str]):
        self.__settings = []
        # find settings
        start_ind, end_ind = -1, -1
        for i in range(len(rows)):
            IS_LAST_ROW = i == (len(rows) - 1)
            HAS_GUID = self._find_index(rows[i], self._SPACE_4_PATTERN) != -1 and self._find_index(rows[i], self._GUID_PATTERN) != -1

            if HAS_GUID and start_ind == -1:
                start_ind = i
            elif start_ind != -1 and HAS_GUID:
                end_ind = i
            elif start_ind != -1 and IS_LAST_ROW:
                end_ind = len(rows)

            if end_ind != -1:
                block = '\n'.join(rows[start_ind:end_ind])
                self.__settings.append(Setting(block))
                start_ind, end_ind = i, -1

    def load_from_json(self, json_data: dict):
        try:
            for setting in self.get_settings():
                setting.load_from_json(json_data['settings'][setting.get_guid()])
        except KeyError:
            pass

    def get_settings(self) -> list[Setting]:
        return self.__settings
    
    def to_json(self) -> Any:
        settings_dict = {}
        for setting in self.get_settings():
            settings_dict[setting.get_guid()] = setting.to_json()

        return {
            'name': self.get_name(),
            'settings': settings_dict
        }

class Scheme(Node):
    __subgroups : list[SubGroup] = None

    def __init__(self, str_to_parse: str) -> None:
        super().__init__(str_to_parse)

    def _parse(self, rows:list[str]):
        self.__subgroups = []
        # find subgroups
        start_ind, end_ind = -1, -1
        for i in range(len(rows)):
            IS_LAST_ROW = i == (len(rows) - 1)
            HAS_GUID = self._find_index(rows[i], self._SPACE_2_PATTERN) != -1 and self._find_index(rows[i], self._GUID_PATTERN) != -1

            if HAS_GUID and start_ind == -1:
                start_ind = i
            elif start_ind != -1 and HAS_GUID:
                end_ind = i
            elif start_ind != -1 and IS_LAST_ROW:
                end_ind = len(rows)

            if end_ind != -1:
                block = '\n'.join(rows[start_ind:end_ind])
                self.__subgroups.append(SubGroup(block))
                start_ind, end_ind = i, -1

    def load_from_json(self, json_data: dict):
        if json_data['guid'] != self.get_guid():
            raise Exception("Wrong guid for schema")
        
        for subgroup in self.get_subgroups():
            try:
                subgroup.load_from_json(json_data['subgroups'][subgroup.get_guid()])
            except KeyError:
                pass

    def get_subgroups(self) -> list[SubGroup]:
        return self.__subgroups
    
    def to_json(self) -> Any:
        subgroups_dict = {}
        for subgroup in self.get_subgroups():
            subgroups_dict[subgroup.get_guid()] = subgroup.to_json()

        return {
            'guid': self.get_guid(),
            'name': self.get_name(),
            'subgroups': subgroups_dict
        }

class PowerCfg:
    __GET_CUR_CFG : str = 'powercfg /query'

    __scheme : Scheme = None

    def __init__(self) -> None:
        self.__scheme = Scheme(PowerCfg.__call_shell(PowerCfg.__GET_CUR_CFG))

    @staticmethod
    def __decode(bytes_obj:bytes) -> str:
        encoding = chardet.detect(bytes_obj)['encoding']
        return bytes_obj.decode(encoding)

    @staticmethod
    def __call_shell(command:str) -> str:
        process = subprocess.run(command, stdout=subprocess.PIPE, shell=True)
        return PowerCfg.__decode(process.stdout)
    
    def get_scheme(self) -> Scheme:
        return self.__scheme
    
    def load_from_json(self, filename:str):
        with open(filename, 'r', encoding='utf-8') as f:
            self.__scheme.load_from_json(json.load(f))
    
    def export_to_json(self, filename:str):
        mode = 'w' if os.path.isfile(filename) else 'x'
        with open(filename, mode, encoding='utf-8') as f:
            json.dump(self.get_scheme().to_json(), f, indent=4, ensure_ascii=False)

    def apply_schema(self):
        for subgroup in self.get_scheme().get_subgroups():
            for setting in subgroup.get_settings():
                if setting.is_ac_changed():
                    command = f'powercfg -setacvalueindex {self.get_scheme().get_guid()} {subgroup.get_guid()} {setting.get_guid()} {setting.get_ac_value_hex()}'
                    subprocess.run(command)
                    print(command)
                if setting.is_dc_changed():
                    command = f'powercfg -setdcvalueindex {self.get_scheme().get_guid()} {subgroup.get_guid()} {setting.get_guid()} {setting.get_dc_value_hex()}'
                    subprocess.run(command)
                    print(command)
                setting.update_old_values()