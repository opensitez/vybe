// vybe-test: go/functions_patterns_extra/variadic_of_interface_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func pack(values ...interface{}) []interface{} { return values }
func main() { _ = pack(1, "two", true) }
