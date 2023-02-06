# Skarb

[![Generic badge](https://img.shields.io/badge/license-GNU_GPL_v.3-COLOR.svg)]()
[![Generic badge](https://img.shields.io/badge/*.exe-pyinstaller-COLOR.svg)]()
[![Generic badge](https://img.shields.io/badge/MS_Word-python_docx-COLOR.svg)]()

Обробка та зведення загальних показників експортованих таблиць доходів ДРФО (J1703502.xsd XML scheme)

![](demo/p_.png)

## Імпорт файлів

Придатні до опрацювання файли формату *.XML (файл формату PDF не опрацьовується), отримані від офіційного розпорядника. 
Власноручне внесення змін до файлу або збереження формату сторонніми програмами може призвести до унеможливлення 
конвертування.

## Експорт

- форматовані таблиці MS Excel
- звіти MS Word:
  
### Загальні та середні суми доходів, джерела доходів:

![](demo/p_1.png)

### Деталізація доходів щорічно:

![](demo/p_2.png)

### Деталізація доходів за видами надходжень:

![](demo/p_3.png)

### Форматована таблиця записів реєстру для друку:

![](demo/p_4.png)
    

## Збірка проекту
1. venv python 3.9 
2. requirements.txt 
3. cd project_dir 
4. pyinstaller --path project_dir\venv\Lib\site-packages --noconfirm --windowed --onefile --icon project_dir/app_icon.ico --name skarb project_dir/main.py

## Особливість опрацювання записів декларування доходів ФОП

Враховуються декларації останнього звітного періоду року (якщо є 6 та 9-місячний звіти - буде враховано тільки 9-місячний). В якості джерела доходу вказується сама особа з власним кодом РНОКПП, до статистики загальної суми доходи зазначається - ДОХОДИ ПІДПРИЄМНИЦЬКОЇ ДІЯЛЬНОСТІ.  

## Credits:
Used in GUI:
- https://www.flaticon.com/free-icons/excel - Excel icons created by Freepik - Flaticon
- https://www.flaticon.com/free-icons/microsoft-word - Microsoft word icons created by Bharat Icons - Flaticon
- https://www.flaticon.com/free-icons/excel - Excel icons created by Bharat Icons - Flaticon