# Systematic Review Screening Tool - Setup & Usage

## Installation

### 1. Install Python Dependencies
```bash
pip install pandas requests biopython lxml beautifulsoup4 openpyxl
```

Or use the requirements file:
```bash
pip install -r requirements.txt
```

### 2. Run the Application
```bash
python main.py
```

## Usage Guide

### Step 1: Search & Retrieve Articles

1. **Configure Email**: Enter your email address (required by PubMed API)
2. **Enter Search Terms**: Use standard PubMed search syntax
3. **Set Max Results**: Choose how many articles to retrieve (default: 100)
4. **Select Databases**: 
   - Check PubMed for automated searching
   - Check Cochrane and browse for a CSV file if you have one
5. **Click "Start Search"**: The tool will search and save results to `search_results.csv`

### Step 2: Configure Keywords (Optional)

1. Go to the **"Screen Articles"** tab
2. Enter inclusion keywords separated by commas (e.g., "randomized, controlled, trial")
3. Click **"Update Keywords"** - these will be highlighted in green when reviewing articles

### Step 3: Screen Articles

1. Review each article one by one
2. Use the **✅ Include** or **❌ Exclude** buttons to make decisions
3. Navigate with **Previous/Next** buttons if needed
4. Your progress is shown at the top of the screen

### Step 4: Export Results

1. Go to the **"Results"** tab
2. View your screening summary and preview
3. Export options:
   - **Export All Results**: Complete dataset with decisions
   - **Export Included Only**: Just the included articles
   - **Export Excluded Only**: Just the excluded articles

## Files Generated

- `search_results.csv`: All retrieved articles
- `included_YYYYMMDD_HHMMSS.csv`: Included articles with decision timestamps
- `excluded_YYYYMMDD_HHMMSS.csv`: Excluded articles with decision timestamps
- `all_results_YYYYMMDD_HHMMSS.csv`: Complete results with all decisions

## PubMed Search Tips

### Basic Search Examples:
- `diabetes AND exercise` - Articles containing both terms
- `"machine learning"` - Exact phrase search
- `cancer OR tumor OR tumour` - Articles with any of these terms
- `hypertension NOT pregnancy` - Hypertension articles excluding pregnancy

### Advanced Search Examples:
- `diabetes[MeSH] AND exercise[MeSH]` - MeSH term search
- `smith[Author]` - Articles by author Smith
- `2020:2023[pdat]` - Articles published 2020-2023
- `randomized controlled trial[pt]` - Only RCTs

### Field Tags:
- `[ti]` - Title
- `[ab]` - Abstract  
- `[au]` - Author
- `[mh]` - MeSH terms
- `[pt]` - Publication type
- `[dp]` - Date of publication

## Cochrane Integration

Since Cochrane doesn't have a public API:

1. **Manual Export**: Go to Cochrane Library, perform your search, export results as CSV
2. **CSV Format**: The tool expects columns like Title, Authors, Abstract, DOI
3. **Browse File**: Use the "Browse" button to select your Cochrane CSV file

## Troubleshooting

### Common Issues:

1. **"Biopython not available" warning**:
   - Install with: `pip install biopython`
   - Tool will use backup method if unavailable

2. **PubMed search fails**:
   - Check your email is entered correctly
   - Verify internet connection
   - Try reducing max results number
   - Check search terms syntax

3. **Cochrane CSV not loading**:
   - Ensure CSV has proper headers (Title, Authors, Abstract, etc.)
   - Check file encoding (should be UTF-8)
   - Make sure file isn't corrupted

4. **Export errors**:
   - Check file permissions in current directory
   - Ensure adequate disk space
   - Close any open CSV files with same names

### Performance Tips:

- **Large datasets**: For >1000 articles, consider screening in batches
- **Keyword highlighting**: Too many keywords may slow down display
- **Memory usage**: Close other applications for very large searches

## Features Summary

✅ **PubMed Integration**: Direct API access with fallback methods  
✅ **Cochrane Support**: CSV import functionality  
✅ **Keyword Highlighting**: Visual emphasis on inclusion terms  
✅ **Swipe-Style Interface**: Efficient one-by-one screening  
✅ **Progress Tracking**: Real-time statistics and progress  
✅ **Multiple Export Options**: Flexible result export formats  
✅ **Timestamped Decisions**: Track when decisions were made  
✅ **Navigation**: Move forward/backward through articles  
✅ **Professional UI**: Clean, intuitive interface

## Data Privacy

- All processing is done locally on your computer
- No data is sent to external servers except for PubMed API calls
- Your screening decisions are stored only in local CSV files
- Email is only used for PubMed API compliance (not stored or transmitted elsewhere)
