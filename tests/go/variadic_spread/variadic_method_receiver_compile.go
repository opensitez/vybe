// vybe-test: go/variadic_spread/variadic_method_receiver_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
type Logger struct{}
func (l Logger) Write(parts ...string) int { return len(parts) }
func main() { _ = Logger{}.Write("a", "b") }
