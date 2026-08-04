// vybe-test: go/declarations_patterns/short_declaration_err_style_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
type simpleErr struct{}
func (simpleErr) Error() string { return "err" }
func pair() (int, error) { return 1, simpleErr{} }
func main() { n, err := pair()
_, _ = n, err }
