// vybe-test: go/embedding_promotion_extended/dual_embedded_fields_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type x struct { a int }
type y struct { b int }
type p struct { x
y }
func main() { var v p
_ = v.a
_ = v.b }
