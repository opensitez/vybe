// vybe-test: go/cover_database_sql/sql_db_exec
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
_, _ = db.Exec("INSERT INTO t VALUES (1)") }
