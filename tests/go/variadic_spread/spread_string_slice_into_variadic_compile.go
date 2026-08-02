// vybe-test: go/variadic_spread/spread_string_slice_into_variadic_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
func take(words ...string) int { return len(words) }
func main() { tail := []string{"x", "y"}
_ = take(tail...) }
