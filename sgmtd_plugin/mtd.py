import os
import json
from . import xlsxwriter
from typing import Any, Optional, List


def dispatch(j, module=None):
    response = None
    t = j.get("$type", "").split(",")[0]

    if t == "Sungero.Metadata.SolutionMetadata":
        response = Solution(j)
    elif t == "Sungero.Metadata.ModuleMetadata":
        response = Module(j)
    elif t == "Sungero.Metadata.LayerModuleMetadata":
        response = LayerModule(j)
    elif t == "Sungero.Metadata.EntityMetadata":
        isCollection = [x for x in j.get("Properties", {}) if x.get("IsReferenceToRootEntity")]
        if isCollection:
            response = Collection(j)
        else:
            response = DataBook(j)
    elif t == "Sungero.Metadata.DocumentMetadata":
        response = Document(j)
    elif t == "Sungero.Metadata.TaskMetadata":
        response = Task(j)
    elif t == "Sungero.Metadata.AssignmentMetadata":
        response = Assignment(j)
    elif t == "Sungero.Metadata.NoticeMetadata":
        response = Notice(j)
    elif t == "Sungero.Metadata.ReportMetadata":
        response = Report(j)

    if response and not isinstance(response, (Solution, Module)):
        response.Module = module

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

    def __init__(self, json_str):
        self.json = {}
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

    @property
    def MtdType(self) -> str:
        """ Имя компоненты. """
        return self.__class__.__name__

    def ExcelHeaders(self) -> List[str]:
        return [self.MtdType, 'NameGuid', 'Name']

    def ExcelData(self) -> List[str]:
        return [self.type, self.NameGuid, self.Name]


class BaseMTD(BasicMTD):
    """Базовый класс для работы с MTD"""
    IsArchive = False
    BaseGuid = ""

    def __init__(self, json_str):
        self.Dependencies = []
        self.Overridden = []
        self.PublicConstants = []
        self.PublicFunctions = []
        self.PublicStructures = []
        self.ResourcesKeys = []
        self.Versions = []
        self.Module = None
        self._parent = None
        super().__init__(json_str)

    def __str__(self):
        return "{}.{}({})".format(self.Module, self.Name, self.NameGuid)

    @property
    def Parent(self):
        if not self._parent:
            self._parent = Singleton().entity.get(self.BaseGuid)
        return self._parent

    @property
    def ParentStr(self):
        if self.Parent:
            return "{}.{}.{}".format(self.Parent.Module.CompanyCode, self.Parent.Module.Name, self.Parent.Name)
        else:
            return "---"


class Action(BasicMTD):
    def __init__(self, item: dict, root_entity):
        self.RootEntity = root_entity
        super().__init__(item)

    def ExcelHeaders(self) -> List[str]:
        return ['Type', 'CompanyCode', 'Module', 'EntityType', 'EntityName', 'Action', 'Guid']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        response = [self.type,
                    root.Module.CompanyCode if root else '---',
                    root.Module.Name if root else '---',
                    root.MtdType if root else '---',
                    root.Name if root else '---',
                    self.Name,
                    self.NameGuid]
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
        return ['Type', 'CompanyCode', 'Module', 'EntityType', 'EntityName', 'ControlName', 'Guid']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        response = [self.type,
                    root.Module.CompanyCode if root else '---',
                    root.Module.Name if root else '---',
                    root.MtdType if root else '---',
                    root.Name if root else '---',
                    self.Name,
                    self.NameGuid]
        return response


class Property(BasicMTD):
    IsAncestorMetadata = False
    IsIdentifier = False
    IsUnique = False
    IsReferenceToRootEntity = False
    EntityGuid = ""

    def __init__(self, item: dict, root_entity):
        self.RootEntity = root_entity
        self.CollectionEntity = None
        super().__init__(item)

    def parse(self):
        super().parse()

    @property
    def FullName(self):
        if self.CollectionEntity:
            return '{} -> {}'.format(self.CollectionEntity.Name, self.Name)
        else:
            return self.Name

    def ExcelHeaders(self) -> List[str]:
        return ['Type', 'CompanyCode', 'Module', 'EntityType', 'EntityName', 'PropertyName', 'PropertyGuid']

    def ExcelData(self) -> List[str]:
        root = self.RootEntity
        response = [self.type,
                    root.Module.CompanyCode if root else '---',
                    root.Module.Name if root else '---',
                    root.MtdType if root else '---',
                    root.Name if root else '---',
                    self.FullName,
                    self.NameGuid]
        return response


class Solution(BaseMTD):
    Version = ""
    CompanyCode = ""

    def parse(self):
        super().parse()

    def __str__(self):
        return "{}.{}".format(self.CompanyCode, self.Name, self.Version)

    def ExcelHeaders(self) -> List[str]:
        return ['Type', 'Version', 'Solution', 'Guid', 'Name', 'PGuid', 'PCompanyCode', 'PName']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,
                    self.Version,
                    '{}.{}'.format(self.CompanyCode, self.Name),
                    self.NameGuid,
                    self.Name,
                    '---',
                    '---',
                    '---']
        return response


class Module(BaseMTD):
    Code = ""
    CompanyCode = ""
    Version = ""
    AssociatedGuid = ""
    LayeredFromGuid = ""
    SolutionGuid = ""
    Override = False

    def __init__(self, json_str):
        self.AsyncHandlers = []
        self.Jobs = []
        self.Cover = []
        self.SpecialFolders = []
        self.Widgets = []
        self._solution = None
        super().__init__(json_str)

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
        return ['Type', 'Version', 'Solution', 'Guid', 'Name', 'PGuid', 'PCompanyCode', 'PName']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,
                    self.Version,
                    str(self.Solution),
                    self.NameGuid,
                    '{}.{}'.format(self.CompanyCode, self.Name),
                    '---',
                    '---',
                    '---']
        return response


class LayerModule(Module):

    def parse(self):
        super().parse()

        if self.AssociatedGuid:
            self.Solution = Singleton().entity.get(self.AssociatedGuid)

    def __str__(self):
        return self.Name

    def ExcelHeaders(self) -> List[str]:
        return ['Type', 'Version', 'Solution', 'Guid', 'Name', 'PGuid', 'PCompanyCode', 'PName']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType,
                    self.Version,
                    str(self.Solution),
                    self.NameGuid,
                    self.Name,
                    self.Parent.NameGuid if self.Parent else '---',
                    self.Parent.CompanyCode if self.Parent else '---',
                    self.Parent.Name if self.Parent else '---']
        return response


class DataBook(BaseMTD):
    AccessRightsMode = ""
    IsAbstract = False
    IsVisible = False

    def __init__(self, json_str):
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
        super().__init__(json_str)

    def __str__(self):
        return "{}.{}.{}".format(self.Module.CompanyCode, self.Module.Name, self.Name)

    def parse(self):
        super().parse()
        for form in self.json.get("Forms", []):
            self.Forms.append(form)
            for control in form.get("Controls"):
                item = Control(control, self)
                self.Controls.append(item)

        for action in self.json.get("Actions", []):
            self.Actions.append(Action(action, self))

        for prop in self.json.get("Properties", []):
            self.Properties.append(Property(prop, self))



    def ExcelHeaders(self) -> List[str]:
        return ['Type', 'CompanyCode', 'Module', 'NameGuid', 'Name', 'ParentGuid', 'ParentCompanyCode', 'ParentName']

    def ExcelData(self) -> List[str]:
        response = [self.MtdType, self.Module.CompanyCode, self.Module.Name, self.NameGuid, self.Name]
        response.append(self.Parent.NameGuid if self.Parent else '---')
        response.append(self.Parent.Module.CompanyCode if self.Parent and self.Parent.Module else '---')
        response.append(self.Parent.Name if self.Parent else '---')
        return response


class Collection(DataBook):
    def __init__(self, json_str):
        self.RootEntity = None
        super().__init__(json_str)


class Document(DataBook):
    pass


class Task(DataBook):
    def __init__(self, json_str):
        self.AttachmentGroups = []
        self.Scheme = {}
        self._root_entity_guid = None
        super().__init__(json_str)


class Assignment(DataBook):
    AssociatedGuid = None

    def __init__(self, json_str):
        self._parent_task = None
        self.AttachmentGroups = []
        self.Scheme = {}
        self._root_entity_guid = None
        super().__init__(json_str)

    @property
    def MainTask(self):
        if not self._parent_task:
            self._parent_task = Singleton().entity.get(self.AssociatedGuid)
        return self._parent_task


class Notice(DataBook):
    pass


class Report(DataBook):
    pass


def parse_file(path, filename, module=None):
    with open(os.path.join(path, filename), 'r', encoding='utf8') as fp:
        response = dispatch(json.load(fp), module)
        if response:
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
            response = parse_file(path, 'Module.mtd')
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
                    response = parse_file(subpath, mtd, module)
                    if not response:
                        print('ERROR', path, subpath, mtd)
                        continue

                    response.IsArchive = is_archive
                    result[response.NameGuid] = response

    # постобработка
    collection_guids = []
    for k in result.keys():
        item = result[k]
        if not isinstance(item, (DataBook)):
            continue

        # обновление родителей после полной загрузки
        if item.Parent == None:
            pass

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
                pc.CollectionEntity = p
                item.Properties.append(pc)

    return result, archive


def render_excel(data, filename):
    wb = xlsxwriter.Workbook(filename)

    header_format = wb.add_format()
    header_format.set_bold()

    # Решения и модули
    sheet = wb.add_worksheet("ModuleSolution")

    rows = [x for x in data if isinstance(x, (Module, Solution))]
    render_excel_sheet(rows, sheet, header_format)

    # Справочники, Документы, Задачи, Задания, Уведомления, Отчеты
    sheet = wb.add_worksheet("Entity")
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
    sheet = wb.add_worksheet("Action")
    render_excel_sheet(actions, sheet, header_format)

    # Свойства
    sheet = wb.add_worksheet("Property")
    render_excel_sheet(properties, sheet, header_format)

    # Контролы
    sheet = wb.add_worksheet("Control")
    render_excel_sheet(controls, sheet, header_format)

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

