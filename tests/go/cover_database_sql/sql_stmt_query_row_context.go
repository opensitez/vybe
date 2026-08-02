// vybe-test: go/cover_database_sql/sql_stmt_query_row_context
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
stmt, _ := db.Prepare("SELECT min(id) FROM t")
if stmt != nil { _ = stmt.QueryRowContext(context.Background()) } }
