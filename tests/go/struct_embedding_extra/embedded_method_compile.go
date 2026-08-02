// vybe-test: go/struct_embedding_extra/embedded_method_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) label() string { return "ok" }
type outer struct { inner }
func main() { var value outer
_ = value.label() }
