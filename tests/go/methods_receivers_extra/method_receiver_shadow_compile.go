// vybe-test: go/methods_receivers_extra/method_receiver_shadow_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) total() int { c := counter{n: c.n + 1}
return c.n }
func main() { _ = counter{}.total() }
