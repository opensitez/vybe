// vybe-test: go/interface_embedding_methods/nil_composite_in_conditional_else_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type doer interface { doWork() }
type actor interface { doer }
func main() { var value actor
if value != nil { value.doWork() } else { _ = 0 } }
