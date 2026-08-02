// vybe-test: go/embedding_promotion_extended/collision_qualifier_left_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type a struct{}
func (a) f() {}
type b struct{}
func (b) f() {}
type c struct { a
b }
func main() { var x c
x.a.f() }
