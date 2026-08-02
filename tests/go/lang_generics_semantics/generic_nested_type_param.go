// vybe-test: go/lang_generics_semantics/generic_nested_type_param
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
type Outer[T any] struct { Inner[T] }
type Inner[U any] struct { V U }
func main() { _ = Outer[int]{} }
