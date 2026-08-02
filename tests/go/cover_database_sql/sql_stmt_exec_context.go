// vybe-test: go/cover_database_sql/sql_stmt_exec_context
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
stmt, _ := db.Prepare("UPDATE t SET v = ?")
if stmt != nil { _, _ = stmt.ExecContext(context.Background(), 2) } }
