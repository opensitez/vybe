// vybe-test: go/embedding_promotion_extended/promoted_method_call_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) f() {}
type outer struct { inner }
func main() { outer{}.f() }
