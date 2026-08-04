// vybe-test: go/scope/blank_identifier_ignore
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
func pair() (int, int) { return 1, 2 } func main() { _, b := pair()
_ = b }
