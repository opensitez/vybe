// vybe-test: go/functions_patterns_extra/named_return_with_defer_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func run() (result int) { defer func() { result++ }()
result = 1
return }
func main() { _ = run }
