// vybe-test: go/cover_database_sql/sql_isolation_level_string
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "database/sql"
func main() { _ = sql.LevelDefault.String()
_ = sql.LevelReadUncommitted.String() }
