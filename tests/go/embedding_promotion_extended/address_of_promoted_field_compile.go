// vybe-test: go/embedding_promotion_extended/address_of_promoted_field_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
type outer struct { inner }
func main() { var o outer
_ = &o.n }
