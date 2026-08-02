// vybe-test: go/functions_patterns_extra/multiple_named_returns_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func split(v int) (left int, right int) { left = v
right = v + 1
return }
func main() { _, _ = split(1) }
