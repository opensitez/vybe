// vybe-test: go/cover_database_sql/sql_tx_options_isolation
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type TxOptions = sql.TxOptions
func main() { opts := TxOptions{Isolation: sql.LevelReadCommitted, ReadOnly: true}
_ = opts.Isolation
_ = opts.ReadOnly }
