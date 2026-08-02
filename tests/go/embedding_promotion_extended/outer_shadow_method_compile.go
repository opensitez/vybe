// vybe-test: go/embedding_promotion_extended/outer_shadow_method_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) m() {}
type outer struct { inner }
func (outer) m() {}
func main() { outer{}.m() }
