// vybe-test: go/cover_database_sql/sql_db_query_context
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
_, _ = db.QueryContext(context.Background(), "SELECT name FROM t WHERE id = ?", 2) }
