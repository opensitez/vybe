// vybe-test: go/struct_tags_reflect/struct_json_rename_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Payload struct { UserID int `json:"user_id"` }
func main() { _ = Payload{} }
