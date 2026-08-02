// vybe-test: go/struct_tags_reflect/struct_tag_on_nested_struct_field_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Inner struct { N int `json:"n"` }
type Outer struct { Child Inner `json:"child"` }
func main() { _ = Outer{} }
