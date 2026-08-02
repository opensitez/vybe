// vybe-test: go/cover_database_sql/sql_null_bool
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NullBool = sql.NullBool
func main() { var nb NullBool
_ = nb.Valid
_ = nb.Bool }
