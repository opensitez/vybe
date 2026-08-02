// vybe-test: go/cover_database_sql/sql_null_int64
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NullInt64 = sql.NullInt64
func main() { var ni NullInt64
_ = ni.Valid
_ = ni.Int64 }
