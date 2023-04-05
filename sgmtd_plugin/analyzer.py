# coding: utf-8
""" Модуль плагина SG MTD Analyzer. """
import os
import os.path
from typing import Any, Optional, List

from components import ui_models
from components.base_component import BaseComponent
from components.component_manager import ComponentManager, component
from py_common.logger import log
from sungero_deploy.instance_service import InstanceService
from sungero_deploy import scripts_config
from common_plugin.yaml_tools import load_yaml_from_file, is_jinja_expression
from jinja2 import Template, StrictUndefined
from . import xlsxwriter
from . import mtd


@component(alias="sgmtd")
class MtdAnalyzer(BaseComponent):
    """ Анализ .mtd файлов разработки """

    def configure_ui_variables(self, ui_variables: List[ui_models.UIVariable]) -> None:
        pass

    def __init__(self, config_path: Optional[str] = None) -> None:
        """
        Конструктор.

        Args:
            config_path: Путь к конфигу.
        """

        super().__init__(config_path)
        self._instance_service = InstanceService(self._tool_name(), scripts_config.get_instance_name(self.config))
        self._component_path = ComponentManager.get_component_folder(self._tool_name())

    def install(self, **kwargs: Any) -> None:
        """
        Установить компоненту.

        Args:
            *kwargs: Словарь произвольных аргументов установки компоненты.
        """
        if not os.path.exists(self._component_path):
            raise RuntimeError(
                f'{self._tool_name()} component is not installed.\n'
                'You can install the component using the following commands:\n'
                f'\tdo components add {self._tool_name()}')

        log.info(f'"{MtdAnalyzer.__name__}" component has been successfully installed.')

    def uninstall(self) -> None:
        """ Удалить компоненту. """
        log.info(f'"{MtdAnalyzer.__name__}" component has been successfully uninstalled.')

    def _tool_name(self) -> str:
        """ Имя компоненты. """
        return self.__class__.__name__

    def _get_mtd_info(self):
        config_dict = load_yaml_from_file(self.config_path)
        services_config = config_dict.get('services_config', {})
        dds_config = services_config.get('DevelopmentStudio', {})

        # рендер переменных из шаблона
        variables_dict = config_dict.get('variables', {})
        git_root_directory = dds_config.get('GIT_ROOT_DIRECTORY')
        if is_jinja_expression(git_root_directory):
            tm = Template(git_root_directory, undefined=StrictUndefined)
            tm.environment.globals['getenv'] = os.getenv
            git_root_directory = tm.render(variables_dict)

        # получение путей до репозиториев
        repositories = dds_config.get("REPOSITORIES", {}).get("repository", {})
        response = []
        archive = []
        for repo in repositories:
            items, arch = mtd.dir_walk(os.path.join(git_root_directory, repo.get('@folderName')))
            # todo: обработка архивов
            response += items.values()
            archive += arch


        # hack - получение родителей
        items, arch = mtd.dir_walk(os.path.join(git_root_directory, '_platform'))
        response += items.values()
        archive += arch

        return response, archive

    def get_mtd_info(self):
        """ MTD. Вывод краткой структуры репозиториев """
        response, archive = self._get_mtd_info()
        return response

    def save_mtd_info(self, filename: str):
        """ MTD. Сохранить данные в Excel. Параметр - имя файла.xlsx """
        items, archive = self._get_mtd_info()
        mtd.render_excel(items, archive, filename)


def init_plugin() -> None:
    """ Инициализировать плагин. """
    try:
        from sungero_plugin.actions import all_services
        from sungero_plugin.build import build
        from sungero_plugin.build_up import build_up
        from sungero_plugin.sungero_components import local_publish_component
        MtdAnalyzer.build = build  # type: ignore
        MtdAnalyzer.local_publish = local_publish_component  # type: ignore
        MtdAnalyzer.build_up = build_up  # type: ignore
        all_services.append(MtdAnalyzer)
    except ImportError:
        pass

init_plugin()
