// vybe-test: go/init_blank_import/init_sets_func_variable_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var twice func(int) int
func init() { twice = func(n int) int { return n * 2 } }
func main() { _ = twice(4) }
