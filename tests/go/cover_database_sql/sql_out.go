// vybe-test: go/cover_database_sql/sql_out
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type Out = sql.Out
func main() { var dest int
_ = Out{Dest: &dest} }
