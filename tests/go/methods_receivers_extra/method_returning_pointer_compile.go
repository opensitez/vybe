// vybe-test: go/methods_receivers_extra/method_returning_pointer_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type node struct { next *node }
func (n node) clone() *node { return &node{} }
func main() { _ = node{}.clone() }
