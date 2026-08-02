// vybe-test: go/cover_database_sql/sql_tx_exec_context
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
tx, _ := db.Begin()
if tx != nil { _, _ = tx.ExecContext(context.Background(), "DELETE FROM t WHERE id = ?", 8) } }
