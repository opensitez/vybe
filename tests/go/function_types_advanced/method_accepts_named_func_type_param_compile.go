// vybe-test: go/function_types_advanced/method_accepts_named_func_type_param_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Predicate func(int) bool
type sieve struct { n int }
func (s sieve) keep(p Predicate) bool { return p(s.n) }
func main() { _ = sieve{n: 2}.keep(Predicate(func(v int) bool { return v%2 == 0 })) }
