// vybe-test: go/embedding_promotion_extended/pointer_receiver_on_value_embed_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
func (i *inner) bump() { i.n++ }
type outer struct { inner }
func main() { var o outer
o.bump() }
