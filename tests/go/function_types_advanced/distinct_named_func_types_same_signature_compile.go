// vybe-test: go/function_types_advanced/distinct_named_func_types_same_signature_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Step func(int) int
type Stage func(int) int
func pipe(v int, s Stage) int { return s(v) }
func main() { var step Step = func(v int) int { return v + 1 }
_ = pipe(1, Stage(step)) }
