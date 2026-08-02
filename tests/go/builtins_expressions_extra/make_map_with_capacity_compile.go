// vybe-test: go/builtins_expressions_extra/make_map_with_capacity_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { values := make(map[string]int, 4)
_ = values }
