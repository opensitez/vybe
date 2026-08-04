// vybe-test: go/composite_literals_extra/nested_anonymous_struct_literal_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
func main() { value := struct { inner struct { n int } }{}
_ = value }
