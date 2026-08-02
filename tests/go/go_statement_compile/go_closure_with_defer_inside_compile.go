// vybe-test: go/go_statement_compile/go_closure_with_defer_inside_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { go func() { defer func() { _ = 1 }()
_ = 0 }() }
