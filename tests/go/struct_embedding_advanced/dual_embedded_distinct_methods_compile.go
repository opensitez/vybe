// vybe-test: go/struct_embedding_advanced/dual_embedded_distinct_methods_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type left struct{}
func (left) side() string { return "L" }
type right struct{}
func (right) edge() string { return "R" }
type pair struct { left
right }
func main() { _ = pair{}.side()
_ = pair{}.edge() }
