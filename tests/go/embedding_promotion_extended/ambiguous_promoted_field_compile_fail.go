// vybe-test: go/embedding_promotion_extended/ambiguous_promoted_field_compile_fail
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile-fail

package main
type a struct { x int }
type b struct { x int }
type c struct { a
b }
func main() { var v c
_ = v.x }
