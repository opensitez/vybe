// vybe-test: go/cover_database_sql/sql_result_rows_affected
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
res, _ := db.Exec("DELETE FROM t")
if res != nil { _, _ = res.RowsAffected() } }
