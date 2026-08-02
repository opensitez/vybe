// vybe-test: go/struct_embedding_advanced/outer_method_shadows_embedded_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) label() string { return "inner" }
type outer struct { inner }
func (outer) label() string { return "outer" }
func main() { _ = outer{}.label() }
