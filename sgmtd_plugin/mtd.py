import os
import json
from codecs import BOM_UTF8
from typing import Any, Optional, List, Dict
import xml.etree.ElementTree as ET

# запуск из разных контекстов
try:
    from . import xlsxwriter
except ImportError:
    import xlsxwriter


def dispatch(mtd_file, module=None, en_file=None, ru_file=None):
    try:
        response = None
        if mtd_file:
            j = json.loads(mtd_file)
        else:
            return response

        en_res = parse_resx(en_file)
        ru_res = parse_resx(ru_file)

        t = j.get("$type", "").split(",")[0]

        if t == "Sungero.Metadata.SolutionMetadata":
            response = Solution(j)
        elif t == "Sungero.Metadata.ModuleMetadata":
            response = Module(j, en_res, ru_res)
        elif t == "Sungero.Metadata.LayerModuleMetadata":
            response = LayerModule(j, en_res, ru_res)
        elif t == "Sungero.Metadata.EntityMetadata":
            isCollection = [x for x in j.get("Properties", {}) if x.get("IsReferenceToRootEntity")]
            if isCollection:
                response = Collection(j, en_res, ru_res)
            else:
                response = DataBook(j, en_res, ru_res)
        elif t == "Sungero.Metadata.DocumentMetadata":
            response = Document(j, en_res, ru_res)
        elif t == "Sungero.Metadata.TaskMetadata":
            response = Task(j, en_res, ru_res)
        elif t == "Sungero.Metadata.AssignmentMetadata":
            response = Assignment(j, en_res, ru_res)
        elif t == "Sungero.Metadata.NoticeMetadata":
            response = Notice(j, en_res, ru_res)
        elif t == "Sungero.Metadata.ReportMetadata":
            response = Report(j, en_res, ru_res)

        if response and not isinstance(response, (Solution, Module)):
            response.Module = module

    except Exception as exc:
        print(exc)

    return response


class Singleton(object):

    def __new__(cls, elem=None):
        if not hasattr(cls, 'instance'):
            cls.instance = super(Singleton, cls).__new__(cls)

        if not hasattr(cls.instance, 'entity'):
            cls.instance.entity = {}

        if not hasattr(cls.instance, 'property'):
            cls.instance.property = {}

        if not hasattr(cls.instance, 'control'):
            cls.instance.control = {}

        return cls.instance


class BasicMTD:
    type = ""
    Name = ""
    NameGuid = ""
    path = ""

    def __init__(self, json_str, en_res=None, ru_res=None):
        self.json = {}
        self.resx = {'en': en_res if en_res else {}, 'ru': ru_res if ru_res else {}}
        if isinstance(json_str, str):
            self.json = json.loads(json_str)
        elif isinstance(json_str, dict):
            self.json = json_str

        self.NameGuid = self.json.get("NameGuid")
        Singleton().entity[self.NameGuid] = self

        self.parse()

    def __str__(self):
        return "{}({})".format(self.Name, self.NameGuid)

    def parse(self):
        for k in [x for x in dir(self) if "__" not in x]:
            v = self.json.get(k)
            if v and not isinstance(v, list):
                setattr(self, k, v)

        self.type = self.json.get("$type", "").split(",")[0]

    def Locale(self, lang):
        """ Возвращает локализованное имя"""
        return None

    @property
    def MtdType(self) -> str:
        """ Имя компоненты. """
        return self.__class__.__name__

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Guid', 'Название', 'Путь']

    def ExcelData(self) -> List[str]:
        return [self.type, self.NameGuid, self.Name, self.path]


class BaseMTD(BasicMTD):
    """Базовый класс для работы с MTD"""
    IsArchive = False
    BaseGuid = ""
    Code = ""

    def __init__(self, json_str, en_res=None, ru_res=None):
        self.Dependencies = []
        self.Overridden = []
        self.PublicConstants = []
        self.PublicFunctions = []
        self.PublicStructures = []
        self.ResourcesKeys = []
        self.Versions = []
        self.Module = None
        self._parent = None
        super().__init__(json_str, en_res, ru_res)

    def __str__(self):
        return "{}.{}({})".format(self.Module, self.Name, self.NameGuid)

    def Locale(self, lang):
        """ Возвращает локализованное имя"""
        response = None
        data = self.resx.get(lang)
        if data:
            response = data.get('DisplayName')
        if not response:
            response = self.Parent.Locale(lang) if self.Parent else None  # TODO: оптимизировать

        return response

    @property
    def Parent(self):
        if not self._parent:
            self._parent = Singleton().entity.get(self.BaseGuid)
        return self._parent

    @property
    def RootParent(self):
        # TODO: наверняка есть возможность оптимизации
        parent = self.Parent
        if parent:
            while parent:
                if parent.Parent:
                    parent = parent.Parent
                else:
                    break
        else:
            parent = self

        return parent

    def FullName(self):
        if isinstance(self.Module, Solution):
            return '{}.{}'.format(self.Module.Name, self.Name)
        elif self.Module:
            module = self.Module.Name
            solution = self.Module.Solution.Name if self.Module.Solution else '-'
        else:
            module = '-'
            solution = '-'
        return '{}.{}.{}'.format(solution, module, self.Name)

    def SQLTable(self):
        return '{}_{}_{}'.format(self.RootParent.Module.CompanyCode if self.Module else '---',
                                 self.RootParent.Module.Code if self.Module else '---',
                                 self.RootParent.Code)


class Action(BasicMTD):
    def __init__(self, item: dict, root_entity):
        self.RootEntity = root_entity
        super().__init__(item)

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Код компании', 'Модуль', 'Тип сущности', 'Название', 'Действие', 'Guid', 'Путь']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        response = [self.type,  # Тип
                    root.Module.CompanyCode if root else '---',  # Код компании
                    root.Module.Name if root else '---',  # Модуль
                    root.MtdType if root else '---',  # Тип сущности
                    root.Name if root else '---',  # Название
                    self.Name,  # Действие
                    self.NameGuid,  # Guid
                    root.path if root else '---'  # Guid
                    ]
        return response


class Control(BasicMTD):
    ParentGuid = ""
    PropertyGuid = ""

    def __init__(self, item: dict, root_entity):
        self._parent = None
        self.RootEntity = root_entity
        Singleton().control[self.NameGuid] = self
        super().__init__(item)

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Код компании', 'Модуль', 'Тип сущности', 'Название сущности', 'Название контрола', 'Guid', 'Путь']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        response = [self.type,
                    root.Module.CompanyCode if root else '---',
                    root.Module.Name if root else '---',
                    root.MtdType if root else '---',
                    root.Name if root else '---',
                    self.Name,
                    self.NameGuid,
                    root.path if root else '---']
        return response


class Property(BasicMTD):
    IsAncestorMetadata = False
    IsIdentifier = False
    IsUnique = False
    IsReferenceToRootEntity = False
    EntityGuid = ""
    Code = ""

    def __init__(self, item: dict, root_entity):
        self.RootEntity = root_entity
        self.CollectionProperty = None
        self.CollectionEntity = None
        super().__init__(item)

    def parse(self):
        super().parse()

    def Locale(self, lang):
        if not self.RootEntity:
            return None
        data = self.RootEntity.resx.get(lang)
        if not data:
            return None
        return data.get('Property_' + self.Name)

    @property
    def FullName(self):
        if self.CollectionProperty:
            return '{} -> {}'.format(self.CollectionProperty.Name, self.Name)
        else:
            return self.Name

    def SQLColumn(self):
        # TODO: некорректно для коллекций
        if self.RootEntity and self.RootEntity.RootParent == self.RootEntity:
            return self.Code if self.Code else self.Name
        elif self.RootEntity:
            return "{}_{}_{}".format(self.Code if self.Code else self.Name,
                                     self.RootEntity.Module.Code,
                                     self.RootEntity.Module.CompanyCode)

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Код компании', 'Модуль', 'Тип сущности', 'Название', 'Свойство', 'Имя[En]', 'Имя[Ru]', 'Guid', 'SQL столбец', 'Путь']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        path = root.path if root else '---'
        if self.CollectionEntity and self.CollectionEntity.path:
            path = self.CollectionEntity.path
        response = [self.type,  # Тип
                    root.Module.CompanyCode if root else '---',  # Код компании'
                    root.Module.Name if root else '---',  # Модуль
                    root.MtdType if root else '---',  # Тип сущности
                    root.Name if root else '---',  # Название сущности
                    self.FullName,  # Свойство
                    self.Locale('en'),  # Имя[En]
                    self.Locale('ru'),  # Имя[Ru]
                    self.NameGuid,  # Guid
                    self.SQLColumn(),  # SQL столбец
                    path  # Путь
                    ]
        return response


class Solution(BaseMTD):
    Version = ""
    CompanyCode = ""

    def parse(self):
        super().parse()

    def __str__(self):
        return "{}.{}".format(self.CompanyCode, self.Name)

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Версия', 'Решение/Модуль', 'Guid', 'Код компании', 'Имя', 'Название[En]', 'Название[Ru]',
                'Guid родителя', 'Код компании родителя', 'Название родителя', 'Путь']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,  # Тип
                    self.Version,  # Версия
                    str(self),  # Решение/Модуль
                    self.NameGuid,  # Guid
                    self.CompanyCode,  # Код компании
                    self.Name,  # Название
                    '---',  # Название[En]
                    '---',  # Название[Ru]
                    '---',  # Код компании родителя
                    '---',  # Guid родителя
                    '---',  # Название родителя
                    self.path  # Путь
                    ]
        return response

    def FullName(self):
        return self.Name


class Module(BaseMTD):
    CompanyCode = ""
    Version = ""
    AssociatedGuid = ""
    LayeredFromGuid = ""
    SolutionGuid = ""
    Override = False

    def __init__(self, json_str, en_res: Dict[str, str], ru_res: Dict[str, str]):
        self.AsyncHandlers = []
        self.Jobs = []
        self.Cover = []
        self.SpecialFolders = []
        self.Widgets = []
        self._solution = None
        super().__init__(json_str, en_res, ru_res)

    def parse(self):
        super().parse()

        # обычный модуль
        deps = [x for x in self.json.get("Dependencies", []) if x and x.get("IsSolutionModule")]
        if deps:
            self.SolutionGuid = deps[0].get("Id")

    def __str__(self):
        return "{}.{}".format(self.CompanyCode, self.Name)

    @property
    def Solution(self):
        if not self._solution:
            self._solution = Singleton().entity.get(self.SolutionGuid)
        return self._solution

    @Solution.setter
    def Solution(self, solution: Solution):
        self._solution = solution

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Версия', 'Решение/Модуль', 'Guid', 'Код компании', 'Имя', 'Название[En]', 'Название[Ru]',
                'Guid родителя', 'Код компании родителя', 'Название родителя', 'Путь']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,  # Тип
                    self.Version,  # Версия
                    str(self.Solution),  # Решение/Модуль
                    self.NameGuid,  # Guid
                    self.CompanyCode,  # Код компании
                    self.Name,  # Имя
                    self.Locale('en'),  # Название[En]
                    self.Locale('ru'),  # Название[Ru]
                    '---',  # Guid родителя
                    '---',  # Код компании родителя
                    '---',  # Название родителя
                    self.path  # Путь
                    ]
        return response

    def FullName(self):
        return '{}.{}'.format(self.Solution.Name if self.Solution else '-',
                              self.Name)


class LayerModule(Module):
    def parse(self):
        super().parse()

        if self.AssociatedGuid:
            self.Solution = Singleton().entity.get(self.AssociatedGuid)

    def __str__(self):
        return self.Name

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Версия', 'Решение/Модуль', 'Guid', 'Код компании', 'Название', 'Имя[En]', 'Имя[Ru]',
                'Guid родителя', 'Код компании родителя', 'Название родителя', 'Путь']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,  # Тип
                    self.Version,  # Версия
                    str(self.Solution),  # Решение/Модуль
                    self.NameGuid,  # Guid
                    self.CompanyCode,  # Код компании
                    self.Name,  # Название
                    self.Locale('en'),  # Имя[En]
                    self.Locale('ru'),  # Имя[Ru]
                    self.Parent.NameGuid if self.Parent else '---',  # Guid родителя
                    self.Parent.CompanyCode if self.Parent else '---',  # Код компании родителя
                    self.Parent.Name if self.Parent else '---',  # Название родителя
                    self.path  # Путь
                    ]
        return response

    def FullName(self):
        return '{}.{}'.format(self.Solution.Name if self.Solution else '-',
                              self.Name)


class DataBook(BaseMTD):
    AccessRightsMode = ""
    IsAbstract = False
    IsVisible = False

    def __init__(self, json_str, en_res=None, ru_res=None):
        self.Actions = []
        self.ConverterFunctions = []
        self.Forms = []
        self.Controls = []
        self.HandledEvents = []
        self.Operations = []
        self.Overridden = []
        self.Properties = []
        self.RibbonCardMetadata = []
        self.RibbonCollectionMetadata = []
        self.Module = None
        super().__init__(json_str, en_res, ru_res)

    def __str__(self):
        return "{}.{}.{}".format(self.Module.CompanyCode, self.Module.Name, self.Name)

    def parse(self):
        super().parse()
        for form in self.json.get("Forms", []):
            self.Forms.append(form)
            for control in form.get("Controls", []):
                item = Control(control, self)
                self.Controls.append(item)

        for action in self.json.get("Actions", []):
            self.Actions.append(Action(action, self))

        for prop in self.json.get("Properties", []):
            self.Properties.append(Property(prop, self))

    def ExcelHeaders(self) -> List[str]:
        return ['Тип', 'Код компании', 'Модуль', 'Guid', 'Название', 'Имя[En]', 'Имя[Ru]', 'SQL таблица',
                'Код компании родителя', 'Guid родителя', 'Название родителя', 'Путь']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,  # Тип
                    self.Module.CompanyCode,  # Код компании
                    self.Module.Name,  # Модуль
                    self.NameGuid,  # Guid
                    self.Name,  # Название
                    self.Locale('en'),  # Имя[En]
                    self.Locale('ru'),  # Имя[Ru]
                    self.SQLTable(),  # SQL таблица
                    self.Parent.Module.CompanyCode if self.Parent and self.Parent.Module else '---',  # Код компании родителя
                    self.Parent.NameGuid if self.Parent else '---',  # Guid родителя
                    self.Parent.Name if self.Parent else '---',  # Название родителя
                    self.path  #
                    ]
        return response


class Collection(DataBook):
    def __init__(self, json_str, en_res=None, ru_res=None):
        self.RootEntity = None
        super().__init__(json_str)


class Document(DataBook):
    def SQLTable(self):
        return 'Sungero_Content_EDoc'


class Task(DataBook):
    def __init__(self, json_str, en_res=None, ru_res=None):
        self.AttachmentGroups = []
        self.Scheme = {}
        self._root_entity_guid = None
        super().__init__(json_str, en_res, ru_res)

    def SQLTable(self):
        return 'Sungero_WF_Task'


class Assignment(DataBook):
    AssociatedGuid = None

    def __init__(self, json_str, en_res=None, ru_res=None):
        self._parent_task = None
        self.AttachmentGroups = []
        self.Scheme = {}
        self._root_entity_guid = None
        super().__init__(json_str, en_res, ru_res)

    @property
    def MainTask(self):
        if not self._parent_task:
            self._parent_task = Singleton().entity.get(self.AssociatedGuid)
        return self._parent_task

    def SQLTable(self):
        return 'Sungero_WF_Assignment'


class Notice(DataBook):
    def SQLTable(self):
        return 'Sungero_WF_Assignment'


class Report(DataBook):
    pass


def get_file(filename: str):
    if not os.path.isfile(filename):
        return None

    with open(filename, 'r', encoding='utf-8-sig') as fp:
        return fp.read()


def parse_resx(resx: str):
    response = {}
    if resx:
        root = ET.fromstring(resx)
        for child in root.iter('data'):
            name = child.get('name')
            value = next(child.iter('value')).text
            response[name] = value
    return response


def parse_file(path, module=None):
    mtd_file = get_file(path)
    if not mtd_file:
        return None

    ru_file = get_file(path.replace('.mtd', 'System.ru.resx'))
    en_file = get_file(path.replace('.mtd', 'System.resx'))

    response = dispatch(mtd_file, module, en_file, ru_file)
    if response:
        response.path = path.replace('/', '\\')
        if 'VersionData' in path:
            response.IsArchive = True

    return response


def dir_walk(repo_path: str):
    result = {}
    archive = []

    module = None
    skip_path = ''
    for path, folders, files in os.walk(repo_path):

        # пропускаем уже обработанные каталоги
        if skip_path and skip_path in path:
            continue

        #  каталог решения / модуля
        is_archive = 'VersionData' in path
        if 'Module.mtd' in files:
            response = parse_file(os.path.join(path, 'Module.mtd'), 'Module.mtd')
            if not response:
                continue

            if isinstance(response, Module) or isinstance(response, Solution):
                module = response
            else:
                print("ERROR", path)

            response.IsArchive = is_archive

            if is_archive:
                archive.append(response)
            else:
                result[response.NameGuid] = response

            skip_path = path

            for folder in folders:
                subpath = os.path.join(path, folder)
                mtds = [x for x in os.listdir(subpath) if '.mtd' in x]
                for mtd in mtds:
                    response = parse_file(os.path.join(subpath, mtd), module)
                    if not response:
                        print('ERROR', path, subpath, mtd)
                        continue

                    response.IsArchive = is_archive

                    if is_archive:
                        archive.append(response)
                    else:
                        result[response.NameGuid] = response

    # постобработка
    for k in result.keys():
        item = result[k]
        if not isinstance(item, BaseMTD):
            continue

        # обновление родителей после полной загрузки
        if item.Parent:
            pass

        if not isinstance(item, DataBook):
            continue

        # подгрузка свойств из коллекций
        for p in [x for x in item.Properties if x.type == 'Sungero.Metadata.CollectionPropertyMetadata']:
            collection = result.get(p.EntityGuid)
            if not collection:
                continue

            for pc in collection.Properties:
                # hack - странная отрисовка ссылки на родителя, заменил на Id, как видится в DDS
                if pc.IsReferenceToRootEntity:
                    pc.Name = 'Id'
                pc.RootEntity = item
                pc.CollectionProperty = p
                pc.CollectionEntity = collection
                item.Properties.append(pc)

    return result, archive


def render_excel(data, archive, filename):
    wb = xlsxwriter.Workbook(filename)

    header_format = wb.add_format()
    header_format.set_bold()

    # Решения и модули
    sheet = wb.add_worksheet("Модули_Решения")

    rows = [x for x in data if isinstance(x, (Module, Solution))]
    render_excel_sheet(rows, sheet, header_format)

    # Справочники, Документы, Задачи, Задания, Уведомления, Отчеты
    sheet = wb.add_worksheet("Сущности")
    rows = [x for x in data if
            isinstance(x, (DataBook, Document, Task, Assignment, Notice, Report)) and not isinstance(x, Collection)]
    render_excel_sheet(rows, sheet, header_format)

    actions = []
    properties = []
    controls = []
    for item in rows:
        for acti in item.Actions:
            actions.append(acti)

        for prop in item.Properties:
            properties.append(prop)

        for cont in item.Controls:
            controls.append(cont)

    # Действия
    sheet = wb.add_worksheet("Действия")
    render_excel_sheet(actions, sheet, header_format)

    # Свойства
    sheet = wb.add_worksheet("Свойства")
    render_excel_sheet(properties, sheet, header_format)

    # Контролы
    sheet = wb.add_worksheet("Контролы")
    render_excel_sheet(controls, sheet, header_format)

    # Архив
    sheet = wb.add_worksheet("Архив")
    archive += [x for x in data if isinstance(x, (Module, Solution, DataBook, Document, Task, Assignment, Notice, Report)) and not isinstance(x, Collection)]

    render_excel_sheet_archive(archive, sheet, header_format)

    wb.close()


def render_excel_sheet(rows: List[BasicMTD], sheet, header_format):
    len_headers = 0
    for row_num, r in enumerate(rows):
        if row_num == 0:
            len_headers = len(r.ExcelHeaders())
            sheet.write_row(0, 0, r.ExcelHeaders(), header_format)

        sheet.write_row(row_num + 1, 0, r.ExcelData())

    if rows and len_headers:
        sheet.autofilter(0, 0, len(rows), len_headers - 1)
        sheet.autofit()


def render_excel_sheet_archive(rows: List[BaseMTD], sheet, header_format):
    len_headers = 0
    headers = ['Type', 'Version', 'Name', 'FullName', 'Guid', 'ParentGuid', 'Path']
    for row_num, r in enumerate(rows):
        if row_num == 0:
            len_headers = len(headers)
            sheet.write_row(0, 0, headers, header_format)

        if isinstance(r, (Module, LayerModule)):
            row = [
                r.type,
                r.Version,
                r.Name,
                r.FullName(),
                r.NameGuid,
                r.Parent.NameGuid if r.Parent else '---',
                r.path
            ]
        else:
            row = [
                r.type,
                r.Module.Version if r.Module else '---',
                r.Name,
                r.FullName(),
                r.NameGuid,
                r.Parent.NameGuid if r.Parent else '---',
                r.path
            ]
        sheet.write_row(row_num + 1, 0, row)

    if rows and len_headers:
        sheet.autofilter(0, 0, len(rows), len_headers - 1)
        sheet.autofit()


if __name__ == "__main__":
    result = []
    result.append(parse_file(r"C:\Storage\git_repository\work\Starkov.Main\Starkov.Main.Shared\Module.mtd"))
    result.append(parse_file(r"c:\Storage\git_repository\base\Sungero.Company\Sungero.Company.Shared\Module.mtd"))
    result.append(parse_file(r"c:\Storage\git_repository\work\Starkov.RMK\Starkov.RMK.Shared\Sungero.Company\Module.mtd"))

    for item in result:
        print(item, item.Locale('en'), item.Locale('ru'))
