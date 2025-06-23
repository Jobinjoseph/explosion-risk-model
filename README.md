# Explosion Risk Estimation Tool (TLM)

This R project calculates the overpressure and impulse duration for explosions in bunker chambers, following formulas from the TLM Luftstoss standard.

## Structure

- `WA_Lager_functions.R` — Core functions for pressure/time calculations.
- `main_script.R` — Orchestrates reading inputs, performing calculations, and writing results.
- `ergebnisse/` — Folder where CSV/XLSX outputs are saved (auto-generated).
- `config_template.R` — Optional configuration structure (do not upload real credentials).

## Inputs

Expected input Excel file with three sheets:
- `Kammern`
- `Stollen`
- `Exposition`

*These input files are not included in the repository for confidentiality reasons.*
For testing, use `data/example_input.xlsx` with 3 sheets:
- Kammern
- Stollen
- Exposition

Do not upload real data.

## Outputs

- Final results in `ergebnisse/`, automatically timestamped.
- Intermediate ratios and logs also saved in `ergebnisse/`.

## Notes

- Do not upload internal data.
- Customize SharePoint/OneDrive connections via environment variables or a config file.
- Tested with R 4.x and packages: `dplyr`, `readxl`, `openxlsx`, `Microsoft365R`, `httr`.

## License

MIT License 
