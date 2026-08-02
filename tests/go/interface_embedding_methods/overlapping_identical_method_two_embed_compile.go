// vybe-test: go/interface_embedding_methods/overlapping_identical_method_two_embed_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type resetterA interface { reset() }
type resetterB interface { reset() }
type dualReset interface { resetterA
resetterB }
type engine struct{}
func (engine) reset() {}
func main() { var value dualReset = engine{}
value.reset() }
