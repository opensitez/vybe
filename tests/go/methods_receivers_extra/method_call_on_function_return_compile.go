// vybe-test: go/methods_receivers_extra/method_call_on_function_return_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) total() int { return c.n }
func build() counter { return counter{} }
func main() { _ = build().total() }
