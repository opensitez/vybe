// vybe-test: go/interface_embedding_methods/nil_named_interface_method_call_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type handler interface { handle() }
type processor interface { handler }
func invoke(value processor) { value.handle() }
func main() { invoke(nil) }
