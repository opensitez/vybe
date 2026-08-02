// vybe-test: go/cover_database_sql/sql_column_type_name
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
rows, _ := db.Query("SELECT 1")
if rows != nil { types, _ := rows.ColumnTypes()
if len(types) > 0 { _, _ = types[0].Name() } } }
