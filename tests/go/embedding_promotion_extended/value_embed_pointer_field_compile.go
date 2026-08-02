// vybe-test: go/embedding_promotion_extended/value_embed_pointer_field_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
type outer struct { *inner }
func main() { o := outer{inner: &inner{n: 1}}
_ = o.n }
