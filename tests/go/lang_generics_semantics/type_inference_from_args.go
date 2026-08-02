// vybe-test: go/lang_generics_semantics/type_inference_from_args
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func Dup[T any](v T) (T,T) { return v, v }
func main() { _, _ = Dup("a") }
