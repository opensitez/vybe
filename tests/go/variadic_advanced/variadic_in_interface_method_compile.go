// vybe-test: go/variadic_advanced/variadic_in_interface_method_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
type Writer interface { Write(parts ...string) int }
type logger struct{}
func (l logger) Write(parts ...string) int { return len(parts) }
func main() { var w Writer = logger{}
_ = w.Write("a") }
