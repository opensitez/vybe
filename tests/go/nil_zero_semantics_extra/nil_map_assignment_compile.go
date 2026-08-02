// vybe-test: go/nil_zero_semantics_extra/nil_map_assignment_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
values["a"] = 1 }
