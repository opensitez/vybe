// vybe-test: go/struct_tags_reflect/struct_json_dash_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Secret struct { Token string `json:"-"` }
func main() { _ = Secret{} }
