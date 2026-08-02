// vybe-test: go/lang_generics_semantics/generic_func_type_param_in_sig
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func Apply[T any, R any](f func(T) R, v T) R { return f(v) }
func main() { _ = Apply(func(int) string { return "" }, 1) }
