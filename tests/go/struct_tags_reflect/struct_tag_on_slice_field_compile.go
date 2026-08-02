// vybe-test: go/struct_tags_reflect/struct_tag_on_slice_field_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Batch struct { Items []string `json:"items"` }
func main() { _ = Batch{} }
