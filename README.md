# Logo Grid PowerPoint Exporter

This Streamlit app allows you to upload or select preloaded logos and automatically generate a PowerPoint slide with them arranged in a grid.

## Setup

1. Clone the repo or download the ZIP.
2. Install dependencies:

```
pip install -r requirements.txt
```

3. Run the app:

```
streamlit run app.py
```

4. Upload logos or place them in the `preloaded_logos` folder before starting.

## Notes

- Logos are resized to fit within a 3:1 ratio box per cell.
- The dropdown menu now supports lazy loading to avoid memory crashes.
- Canvas grid size can be customized in inches.
