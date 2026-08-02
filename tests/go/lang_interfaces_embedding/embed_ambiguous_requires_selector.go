// vybe-test: go/lang_interfaces_embedding/embed_ambiguous_requires_selector
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type A struct{}
func (A) F() {}
type B struct{}
func (B) F() {}
type C struct { A
B }
func main() { var c C
_ = c }
