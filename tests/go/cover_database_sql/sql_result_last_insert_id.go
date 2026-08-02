// vybe-test: go/cover_database_sql/sql_result_last_insert_id
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
res, _ := db.Exec("INSERT INTO t VALUES (1)")
if res != nil { _, _ = res.LastInsertId() } }
