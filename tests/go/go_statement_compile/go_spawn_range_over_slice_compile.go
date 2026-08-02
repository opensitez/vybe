// vybe-test: go/go_statement_compile/go_spawn_range_over_slice_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { for _, v := range []int{1, 2} { go func() { _ = v }() } }
