// vybe-test: go/methods_receivers_extra/method_with_named_return_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) total() (result int) { result = c.n
return }
func main() { _ = counter{}.total() }
