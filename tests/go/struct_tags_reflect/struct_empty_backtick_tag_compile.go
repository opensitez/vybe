// vybe-test: go/struct_tags_reflect/struct_empty_backtick_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Plain struct { Value int `` }
func main() { _ = Plain{} }
