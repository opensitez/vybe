// vybe-test: go/interface_embedding_methods/overlapping_with_pointer_receiver_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type flushA interface { flush() }
type flushB interface { flush() }
type sink interface { flushA
flushB }
type pipe struct{}
func (p *pipe) flush() {}
func main() { var s sink = &pipe{}
s.flush() }
