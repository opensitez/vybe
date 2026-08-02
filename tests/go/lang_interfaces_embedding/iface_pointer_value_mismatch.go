// vybe-test: go/lang_interfaces_embedding/iface_pointer_value_mismatch
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type I interface { M() }
type T int
func (T) M() {}
func main() { var _ I = T(0) }
