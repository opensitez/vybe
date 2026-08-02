// vybe-test: go/blank_identifier_extended/blank_discard_unary_operators_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { x := 3
_ = -x
_ = ^x }
