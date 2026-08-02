// vybe-test: go/struct_tags_reflect/struct_db_column_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type User struct { Email string `db:"email_addr"` }
func main() { _ = User{} }
