// vybe-test: go/go_statement_compile/go_closure_nested_spawn_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { go func() { go func() { _ = 1 }() }() }
