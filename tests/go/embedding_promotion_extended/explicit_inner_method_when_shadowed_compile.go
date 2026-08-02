// vybe-test: go/embedding_promotion_extended/explicit_inner_method_when_shadowed_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) m() {}
type outer struct { inner }
func (outer) m() {}
func main() { var o outer
o.inner.m() }
