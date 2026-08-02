// vybe-test: go/functions_patterns_extra/named_result_with_short_decl_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func run() (result int) { if value := 3; value > 0 { result = value }
return }
func main() { _ = run }
