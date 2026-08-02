// vybe-test: go/functions_patterns_extra/variadic_parameter_without_use_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func log(values ...int) {}
func main() { log() }
