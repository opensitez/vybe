// vybe-test: go/cover_database_sql/sql_null_string
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NullString = sql.NullString
func main() { var ns NullString
_ = ns.Valid
_ = ns.String }
