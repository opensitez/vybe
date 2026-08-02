// vybe-test: go/cover_database_sql/sql_null_float64
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type NullFloat64 = sql.NullFloat64
func main() { var nf NullFloat64
_ = nf.Valid
_ = nf.Float64 }
