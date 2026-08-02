// vybe-test: go/interface_embedding_methods/nil_composite_interface_unchecked_call_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type worker interface { work() }
type job interface { worker }
func main() { var value job
value.work() }
