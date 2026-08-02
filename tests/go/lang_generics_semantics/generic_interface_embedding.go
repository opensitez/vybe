// vybe-test: go/lang_generics_semantics/generic_interface_embedding
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
type I[T any] interface { Get() T }
type J[T any] interface { I[T]
Set(T) }
func main() { var _ J[int] }
