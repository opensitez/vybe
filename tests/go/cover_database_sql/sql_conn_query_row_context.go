// vybe-test: go/cover_database_sql/sql_conn_query_row_context
// origin: languages/go/tests/go/test_cover_database_sql.rs
// vybe-test-mode: compile

package main
import "context"
import "database/sql"
func main() { db, _ := sql.Open("stub", ":memory:")
defer db.Close()
conn, _ := db.Conn(context.Background())
if conn != nil { _ = conn.QueryRowContext(context.Background(), "SELECT 2") } }
