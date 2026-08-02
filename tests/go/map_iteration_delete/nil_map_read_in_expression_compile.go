// vybe-test: go/map_iteration_delete/nil_map_read_in_expression_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
_ = values["k"] + 1 }
