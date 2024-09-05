# MTDAnalyze

Для более подробного знакомства с утилитой смотрите [документацию с примерами использования.](docs/README.md)



Плагин для DirectumLauncher, который позволяет проанализировать метаданные разработки.

**Процедура установки:** 
Скопировать архив с плагином в каталог DirectumLauncher и выполнить команду:  
`do.bat components add sgmtd`  

**Внимание**: командная строка должна быть запущена с правами администратора и перед запуском установки надо перейти в каталог DirectumLauncher, пример:  

```cmd
c:
cd c:\DirectumLauncher
```

## Использование:

**Генерация отчета**:  
`do.bat sgmtd save_mtd_info ИМЯ_ФАЙЛА.xlsx`

**Генерация файла для автосборки**:
`do.bat sgmtd gen_package package.xml`  



После запуска производится чтение и анализ файлов репозиториев, указанных в `config.yml`, результат сохраняется в указанный Excel файл.
В файле содержится сводная информация обо всех сущностях. В данный момент разбито на вкладки:    

- Решения и Модули
- Сущности (справочники, документы, задачи, задания, уведомления, отчеты), с информацией о перекрытии
- Действия сущностей
- Свойства сущностей
- Контролы



**Процедура удаления:**  

1. Удалить плагин
   
   ```cmd
   do.bat components delete sgmtd
   ```
2. Удалить архив с плагином из каталога DirectumLauncher

> [!NOTE]
> Замечания и пожеланию по развитию шаблона разработки фиксируйте через [Issues](https://github.com/STARKOV-Group/MTDAnalyze/issues).  
> При оформлении ошибки, опишите сценарий для воспроизведения. Для пожеланий приведите обоснование для описываемых изменений - частоту использования, бизнес-ценность, риски и/или эффект от реализации.  
> Внимание! Изменения будут вноситься только в новые версии.  
