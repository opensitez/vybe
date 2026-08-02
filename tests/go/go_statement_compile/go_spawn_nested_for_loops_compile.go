// vybe-test: go/go_statement_compile/go_spawn_nested_for_loops_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { for i := 0; i < 2; i++ { for j := 0; j < 2; j++ { go func() { _, _ = i, j }() } } }
