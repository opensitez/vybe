// vybe-test: go/go_statement_compile/go_spawn_for_loop_index_capture_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { for i := 0; i < 3; i++ { go func() { _ = i }() } }
