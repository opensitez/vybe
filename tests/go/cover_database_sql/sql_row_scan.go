// vybe-test: go/cover_database_sql/sql_row_scan
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
row := db.QueryRow("SELECT 1")
var n int
_ = row.Scan(&n) }
