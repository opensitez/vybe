// vybe-test: go/lang_generics_semantics/generic_pointer_receiver
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
type P[T any] struct { v T }
func (p *P[T]) Set(v T) { p.v = v }
func main() { p := &P[int]{}
p.Set(1) }
