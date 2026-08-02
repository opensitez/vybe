// vybe-test: go/embedding_promotion_extended/embed_value_vs_pointer_type_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type cell struct { n int }
type byValue struct { cell }
type byPtr struct { *cell }
func main() { _ = byValue{}
_ = byPtr{} }
