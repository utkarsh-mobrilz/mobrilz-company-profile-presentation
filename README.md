# Mobrilz Private Limited — Company Profile Presentation

Ultra-premium, modern 25-slide corporate company profile presentation for **Mobrilz Private Limited** ([mobrilz.com](https://www.mobrilz.com/)).

## 📦 Deliverables Included

1. **PowerPoint Presentation File:**
   - [`Mobrilz_Company_Profile_Presentation_v2.pptx`](./Mobrilz_Company_Profile_Presentation_v2.pptx): Fully formatted, editable 16:9 widescreen PowerPoint deck with high-resolution brand assets, balanced typography, custom vector graphics, and clean card styling.
   - Built via `build_light_v2.py`.

2. **Interactive HTML Slide Deck Prototype:**
   - [`presentation_prototype.html`](./presentation_prototype.html): Single-page interactive prototype with keyboard navigation (`ArrowRight`/`ArrowLeft`, `Space`), bottom thumbnail carousel, and dropdown selector.
   - Also configured for GitHub Pages hosting (`index.html` duplicate provided).

3. **Production Brand Assets:**
   - `processed_assets/`: Transparent device mockups (TaxiScout24, Living Art, Mindto, Ayrton Coffee, SF50 Flight Support), client brand logos (Mindto, Tula Family, LocaFri, Niyyah Marriage, KoreMe, Chalo-Reisen), executive portrait assets, and subtle wave background watermarks.

---

## 🚀 Running the Interactive Deck Locally

Open `presentation_prototype.html` directly in any modern browser, or launch a local lightweight server:

```bash
python -m http.server 8080
```
Then navigate to: `http://localhost:8080/presentation_prototype.html`

---

## 🛠 Recompiling the PowerPoint Presentation

Ensure `python-pptx` and `Pillow` are installed:

```bash
pip install python-pptx Pillow
python build_light_v2.py
```

Generated file: `Mobrilz_Company_Profile_Presentation_v2.pptx`
