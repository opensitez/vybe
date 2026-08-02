// vybe-test: go/cover_database_sql/sql_stmt_close
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
stmt, _ := db.Prepare("SELECT 1")
if stmt != nil { _ = stmt.Close() } }
