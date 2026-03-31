# SQLite Data Binding Demo

This example demonstrates how to use data-bound controls with an SQLite database.

## Features

- **SqlConnection** – connects to a local SQLite database (`contacts.db`)
- **SqlDataAdapter** – executes a SELECT query and fills a DataTable
- **DataTable** – holds the result set in memory
- **BindingSource** – acts as the intermediary between the DataTable and the UI controls
- **DataGridView** – displays all contacts, bound through the BindingSource
- **TextBox DataBindings** – individual text fields bound to Name, Email, Phone columns
- **CRUD operations** – Add, Delete, and Refresh buttons

## Data Flow

```
SQLite DB
   ↓
SqlConnection("Data Source=contacts.db")
   ↓
SqlDataAdapter("SELECT * FROM contacts", conn)
   ↓  .Fill(dt)
DataTable
   ↓
BindingSource.DataSource = dt
   ↓                    ↓
DataGridView      TextBox.DataBindings.Add("Text", bs, "ColumnName")
```

## Running

```bash
cargo run --release -- examples/sqlite_binding/SqliteBinding.vbproj
```

The app will create `contacts.db` in the working directory with sample data on first run.
