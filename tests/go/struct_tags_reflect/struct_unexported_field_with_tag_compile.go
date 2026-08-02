// vybe-test: go/struct_tags_reflect/struct_unexported_field_with_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type cache struct { hits int `json:"hits"` }
func main() { _ = cache{} }
