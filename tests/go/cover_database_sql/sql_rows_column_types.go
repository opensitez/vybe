// vybe-test: go/cover_database_sql/sql_rows_column_types
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
rows, _ := db.Query("SELECT 1")
if rows != nil { _, _ = rows.ColumnTypes() } }
