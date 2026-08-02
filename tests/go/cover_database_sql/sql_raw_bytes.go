// vybe-test: go/cover_database_sql/sql_raw_bytes
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
type RawBytes = sql.RawBytes
func main() { var rb RawBytes
_ = rb }
