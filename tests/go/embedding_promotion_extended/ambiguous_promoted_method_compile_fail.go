// vybe-test: go/embedding_promotion_extended/ambiguous_promoted_method_compile_fail
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile-fail

package main
type a struct{}
func (a) f() {}
type b struct{}
func (b) f() {}
type c struct { a
b }
func main() { var v c
v.f() }
