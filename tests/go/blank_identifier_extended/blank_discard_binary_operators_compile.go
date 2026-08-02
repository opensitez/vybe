// vybe-test: go/blank_identifier_extended/blank_discard_binary_operators_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { _, _ = 1+2, 3*4 }
