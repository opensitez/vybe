// vybe-test: go/interfaces_patterns_extra/interface_in_variadic_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func pack(values ...interface{}) []interface{} { return values }
func main() { _ = pack(1, "two") }
