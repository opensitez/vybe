// vybe-test: go/struct_tags_reflect/struct_tag_on_embedded_anonymous_field_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Meta struct { Version string `json:"version"` }
type Doc struct { Meta
Body string `json:"body"` }
func main() { _ = Doc{} }
