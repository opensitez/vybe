// vybe-test: go/lang_interfaces_embedding/interface_method_pointer_receiver_only
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type I interface { M() }
type T int
func (t *T) M() {}
func main() { var _ I = (*T)(nil) }
