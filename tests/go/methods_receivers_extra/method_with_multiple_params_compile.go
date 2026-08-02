// vybe-test: go/methods_receivers_extra/method_with_multiple_params_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) add(a int, b int) int { return c.n + a + b }
func main() { _ = counter{}.add(1, 2) }
