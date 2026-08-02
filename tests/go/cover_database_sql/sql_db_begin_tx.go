// vybe-test: go/cover_database_sql/sql_db_begin_tx
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
_, _ = db.BeginTx(context.Background(), &sql.TxOptions{Isolation: sql.LevelSerializable}) }
