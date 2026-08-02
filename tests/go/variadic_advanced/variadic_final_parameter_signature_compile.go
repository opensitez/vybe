// vybe-test: go/variadic_advanced/variadic_final_parameter_signature_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func take(a int, rest ...string) int { return a + len(rest) }
func main() { _ = take(1, "x", "y") }
