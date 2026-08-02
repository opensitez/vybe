// vybe-test: go/variadic_advanced/variadic_interface_empty_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func pack(values ...interface{}) int { return len(values) }
func main() { _ = pack() }
