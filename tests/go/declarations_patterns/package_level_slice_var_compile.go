// vybe-test: go/declarations_patterns/package_level_slice_var_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
var ids = []int{1, 2, 3}
func main() { _ = ids }
