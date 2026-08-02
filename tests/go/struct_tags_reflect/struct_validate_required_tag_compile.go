// vybe-test: go/struct_tags_reflect/struct_validate_required_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Form struct { Age int `validate:"required,min=0"` }
func main() { _ = Form{} }
