// vybe-test: go/interface_embedding_methods/composite_satisfies_via_struct_literal_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type lock interface { lock() }
type unlock interface { unlock() }
type mutex interface { lock
unlock }
type gate struct{}
func (gate) lock() {}
func (gate) unlock() {}
func main() { var m mutex = gate{}
m.lock() }
