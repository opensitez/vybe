// vybe-test: go/methods_receivers_extra/method_with_bool_return_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type gate struct { open bool }
func (g gate) ready() bool { return g.open }
func main() { _ = gate{}.ready() }
