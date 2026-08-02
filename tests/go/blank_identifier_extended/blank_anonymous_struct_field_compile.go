// vybe-test: go/blank_identifier_extended/blank_anonymous_struct_field_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { type inner struct { _ int
v int }
x := inner{v: 3}
_ = x.v }
