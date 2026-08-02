// vybe-test: go/variadic_spread/mixed_params_int_string_variadic_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
func log(level int, tag string, msgs ...string) int { return level + len(tag) + len(msgs) }
func main() { _ = log(2, "app", "a", "b") }
