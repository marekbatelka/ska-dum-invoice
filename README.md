Nastroj na generovani faktur 

work in progress - silne hardcodovano, zatim prototyp

## Setup

1. Install dependencies: `pipenv sync`
2. Copy `credentials.json.example` to `credentials.json` and fill in your Google API credentials
3. Update `config.py` with your spreadsheet IDs
4. Run: `pipenv run python main.py`

## Changelog
[0.0.1] - 2025-11-17
- vstup google spreadsheet (formular rezervaci)
- vystup google spreadsheet (sablona faktury)
- prida dalsi list a prepise nekolik policek
- umi vygenerovat QR kod (zatim nutno manualne vlozit)
