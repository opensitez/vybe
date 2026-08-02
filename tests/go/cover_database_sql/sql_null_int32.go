// vybe-test: go/cover_database_sql/sql_null_int32
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NullInt32 = sql.NullInt32
func main() { var ni NullInt32
_ = ni.Valid
_ = ni.Int32 }
