// vybe-test: go/blank_identifier_extended/blank_discard_composite_map_literal_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { _ = map[string]int{"k": 1}
m := map[string]int{}
_ = m }
