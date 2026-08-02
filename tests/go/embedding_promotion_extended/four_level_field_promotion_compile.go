// vybe-test: go/embedding_promotion_extended/four_level_field_promotion_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type d struct { n int }
type c struct { d }
type b struct { c }
type a struct { b }
func main() { _ = a{}.n }
