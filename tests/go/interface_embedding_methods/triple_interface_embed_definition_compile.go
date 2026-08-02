// vybe-test: go/interface_embedding_methods/triple_interface_embed_definition_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type leaf interface { tag() string }
type branch interface { leaf }
type trunk interface { branch }
func main() {}
