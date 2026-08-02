// vybe-test: go/cover_database_sql/sql_null_time
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
import "time"
type NullTime = sql.NullTime
func main() { var nt NullTime
_ = nt.Valid
_ = nt.Time
_ = time.Now() }
