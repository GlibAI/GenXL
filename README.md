# GenXL

A tool that converts structured JSON data into professionally formatted Excel reports using LLM-powered layout generation.

## How It Works

1. **Input**: Takes extracted document data as JSON (e.g., bank statements, invoices)
2. **LLM Processing**: Sends the data to Gemini API, which determines optimal Excel cell layout and styling
3. **Excel Generation**: Produces a styled `.xlsx` file with proper formatting, borders, and color coding

## Setup

### Prerequisites

- Python 3.10+
- [Poetry](https://python-poetry.org/) for dependency management
- A [Gemini API key](https://aistudio.google.com/apikey)

### Installation

```bash
poetry install
```

### Configuration

Create a `.env` file in the project root:

```
GEMINI_API_KEY=your_api_key_here
```

## Usage

Place your input JSON file in the project directory and run:

```bash
poetry run python main.py
```

By default, it reads from `testing_jsons/testing_json_1.json` and outputs to `output.xlsx`.

### Input Format

The input JSON should contain extracted document data with the following structure:

```json
{
  "file_name": "document.pdf",
  "classified_file_type": "bank_statement",
  "fields": [
    {
      "field_name": "Account Holder",
      "field_key": "account_holder",
      "section": "Account Information",
      "data_type": "String",
      "value": "John Doe"
    },
    {
      "field_name": "Account Balance",
      "field_key": "account_balance",
      "section": "Account Information",
      "data_type": "Number",
      "value": 15250.75
    },
    {
      "field_name": "Statement Date",
      "field_key": "statement_date",
      "section": "Account Information",
      "data_type": "Date",
      "value": "2024-01-15"
    },
    {
      "field_name": "Transactions",
      "field_key": "transactions",
      "section": "Transaction History",
      "data_type": "Table",
      "value": [
        { "Date": "2024-01-01", "Description": "Direct Deposit", "Amount": 3000.00 },
        { "Date": "2024-01-05", "Description": "Grocery Store", "Amount": -85.50 }
      ]
    }
  ]
}
```

### Field Data Types

| Data Type | Description | Example Value |
|-----------|-------------|---------------|
| `String`  | Plain text values | `"John Doe"` |
| `Number`  | Numeric values (integers or decimals) | `1500.75` |
| `Date`    | Date strings in any format | `"2024-01-15"` |
| `Table`   | Array of objects representing tabular data | `[{"col1": "val1", "col2": "val2"}]` |

### Field Schema

Each field object in the `fields` array supports:

| Key | Type | Description |
|-----|------|-------------|
| `field_name` | string | Human-readable label displayed in the Excel output |
| `field_key` | string | Machine identifier for the field |
| `section` | string | Logical grouping (e.g., "Account Information", "Transaction History") |
| `data_type` | string | One of: `String`, `Number`, `Date`, `Table` |
| `value` | any | The extracted value — can be `null`, a scalar, or an array of objects (for `Table` type) |

## Dependencies

- **openpyxl** - Excel file generation and styling
- **google-genai** - Gemini API client
- **py-toon-format** - Compact data encoding for LLM prompts
- **python-dotenv** - Environment variable management

## License

See [LICENSE](LICENSE) for details.
