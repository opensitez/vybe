// vybe-test: go/cover_database_sql/sql_tx_prepare
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
tx, _ := db.Begin()
if tx != nil { _, _ = tx.Prepare("UPDATE t SET v = ? WHERE id = ?") } }
