// vybe-test: go/struct_embedding_advanced/deep_method_promotion_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type leaf struct{}
func (leaf) tag() string { return "deep" }
type branch struct { leaf }
type trunk struct { branch }
func main() { _ = trunk{}.tag() }
