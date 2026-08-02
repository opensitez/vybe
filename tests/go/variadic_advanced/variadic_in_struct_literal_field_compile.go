// vybe-test: go/variadic_advanced/variadic_in_struct_literal_field_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
type cfg struct { tags []string }
func labels(parts ...string) cfg { return cfg{tags: parts} }
func main() { c := labels("a", "b")
_ = c.tags[1] }
