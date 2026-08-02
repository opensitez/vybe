// vybe-test: go/go_statement_compile/go_promoted_embedded_method_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) work() {}
type outer struct { inner }
func main() { o := outer{}
go o.work() }
