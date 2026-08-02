// vybe-test: go/cover_database_sql/sql_db_stats_fields
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
s := db.Stats()
_ = s.MaxOpenConnections
_ = s.OpenConnections
_ = s.InUse
_ = s.Idle }
