// vybe-test: go/lang_interfaces_embedding/embedding_interface_field
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type I interface { F() }
type S struct { I }
func main() { _ = S{} }
