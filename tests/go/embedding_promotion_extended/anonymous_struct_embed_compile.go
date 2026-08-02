// vybe-test: go/embedding_promotion_extended/anonymous_struct_embed_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type outer struct { struct { x int } }
func main() { _ = outer{}.x }
