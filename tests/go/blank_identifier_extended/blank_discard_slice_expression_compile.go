// vybe-test: go/blank_identifier_extended/blank_discard_slice_expression_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { s := []int{1, 2, 3}
_ = s[1:2]
_ = s[0] }
