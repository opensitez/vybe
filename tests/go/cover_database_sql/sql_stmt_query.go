// vybe-test: go/cover_database_sql/sql_stmt_query
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
stmt, _ := db.Prepare("SELECT id FROM t WHERE id = ?")
if stmt != nil { _, _ = stmt.Query(3) } }
