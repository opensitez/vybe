// vybe-test: go/embedding_promotion_extended/nested_pointer_middle_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
type middle struct { *inner }
type outer struct { middle }
func main() { var o outer
_ = o.n }
