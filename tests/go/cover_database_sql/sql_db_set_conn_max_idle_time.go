// vybe-test: go/cover_database_sql/sql_db_set_conn_max_idle_time
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
import "time"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
db.SetConnMaxIdleTime(time.Minute) }
