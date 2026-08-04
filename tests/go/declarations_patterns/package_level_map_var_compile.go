// vybe-test: go/declarations_patterns/package_level_map_var_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
var lookup = map[string]int{"go": 1}
func main() { _ = lookup }
