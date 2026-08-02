// vybe-test: go/cover_database_sql/sql_tx_stmt
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
stmt, _ := db.Prepare("SELECT 1")
tx, _ := db.Begin()
if tx != nil && stmt != nil { _ = tx.Stmt(stmt) } }
