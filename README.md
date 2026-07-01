# FIFA Team Statistics Scraper

Script en Python para extraer estadisticas de equipos desde la pagina de FIFA
con Selenium. El script hace click en cada boton de categoria, espera a que se
actualice el contenedor `.top-performer-group_mainContent__HbvU9`, limpia los
datos y crea un archivo Excel por cada categoria.

## Instalacion

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Uso

```powershell
python fifa_team_statistics.py
```

La salida queda en:

```text
outputs/fifa_team_statistics/excel_por_categoria/
```

Archivos generados:

- `Attacking.xlsx`
- `Distribution.xlsx`
- `Defending.xlsx`
- `Discipline.xlsx`
- `Goalkeeping.xlsx`
- `Movement.xlsx`
- `Physical.xlsx`

Cada archivo contiene:

- `Datos limpios`: tabla limpia de esa categoria.
- `Resumen numerico`: resumen de columnas numericas, si existen.

## Opciones utiles

Abrir el navegador visible:

```powershell
python fifa_team_statistics.py --headed
```

Esperar mas tiempo despues de cada click:

```powershell
python fifa_team_statistics.py --click-wait-sec 5
```

Consultar solo algunas categorias:

```powershell
python fifa_team_statistics.py --categories Attacking Distribution Defending
```

Cambiar la carpeta de salida:

```powershell
python fifa_team_statistics.py --out-dir outputs/fifa
```

Usar otra URL:

```powershell
python fifa_team_statistics.py --url "https://www.fifa.com/en/tournaments/mens/worldcup/canadamexicousa2026/statistics/team-statistics"
```
