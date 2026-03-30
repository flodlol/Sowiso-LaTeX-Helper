# Sowiso LaTeX Helper

Kleine Chrome-extensie die ik heb gemaakt om sneller wiskunde in Sowiso in te voeren.

Taal:
- Nederlands (dit bestand)
- [English](README.md)

## Wat Het Doet

- Opent als side panel in Chrome.
- Laat je LaTeX plakken of typen.
- Toont een live preview.
- Zet de formule in het actieve Sowiso-invoerveld.
- Ondersteunt 3 invoermodi:
  - `Raw LaTeX`
  - `Inline ($...$)`
  - `Block ($$...$$)`
- Optionele conversie naar Sowiso linear input.
- Thema-switcher: `Auto`, `Light`, `Dark`.
- Bevat een debuglog (met copy/clear knoppen) voor snelle troubleshooting.

## Ondersteunde Pagina's

- `https://cloud.sowiso.nl/*`
- `https://*.sowiso.nl/*`

## Installatie (Unpacked)

1. Open Chrome en ga naar `chrome://extensions`.
2. Zet `Developer mode` aan (rechtsboven).
3. Klik op `Load unpacked`.
4. Selecteer: `chrome-extension`
5. (Optioneel) pin de extensie in je toolbar.

## Gebruik

1. Open een Sowiso-opgavenpagina.
2. Klik een keer in het antwoordveld waar je wilt invoeren.
3. Open het side panel van de extensie.
4. Vul je LaTeX in.
5. Kies de invoermodus.
6. (Optioneel) zet `Convert LaTeX to Sowiso linear input` aan.
7. Klik op `Insert into page` of klik op het preview-blok.

## Opmerkingen

- Preview rendering gebruikt CodeCogs (`https://latex.codecogs.com`) en heeft internet nodig.
- Als invoegen mislukt, klik opnieuw in het antwoordveld en probeer het nog een keer.
- De extensie probeert meerdere invoegmethodes en frame-contexten.

## Privacy

- Geen backend-server voor deze extensie.
- Instellingen worden lokaal opgeslagen via `chrome.storage.local` (o.a. themamodus).
- Je formule wordt lokaal gebruikt voor invoegen en alleen naar CodeCogs gestuurd voor de preview-afbeelding.

## Development

Projectstructuur:
- `chrome-extension/` - broncode van de extensie (manifest, popup, background, content script)
- `public/` - logo's en design-assets

Lokaal draaien:
1. Laad `chrome-extension` als unpacked extensie.
2. Maak codewijzigingen.
3. Herlaad de extensie in `chrome://extensions`.

## Links

- Sponsor: [Study-Track](https://study-track.app/?ref=sowiso-latex-helper)
- GitHub: [flodlol/Sowiso-LaTeX-Helper](https://github.com/flodlol/Sowiso-LaTeX-Helper)

## Licentie

MIT. Zie [`LICENSE`](LICENSE).
