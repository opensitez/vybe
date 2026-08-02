// vybe-test: go/cover_database_sql/sql_named_arg
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NamedArg = sql.NamedArg
func main() { _ = sql.Named("id", 1)
_ = NamedArg{Name: "name", Value: "go"} }
