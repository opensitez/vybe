// vybe-test: go/struct_tags_reflect/struct_tag_on_bool_field_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Flags struct { Active bool `json:"active"` }
func main() { _ = Flags{} }
