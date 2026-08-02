// vybe-test: go/interface_embedding_methods/composite_interface_method_expression_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type mover interface { move() int }
type walker interface { mover }
type step func() int
func (s step) move() int { return s() }
func main() { var fn func(walker) func() int = walker.move
_ = fn }
