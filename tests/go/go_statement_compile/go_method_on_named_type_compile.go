// vybe-test: go/go_statement_compile/go_method_on_named_type_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type counter int
func (c counter) inc() {}
func main() { var c counter
go c.inc() }
