// vybe-test: go/functions_patterns_extra/closure_over_loop_var_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { fns := []func() int{}
for i := 0; i < 2; i++ { fns = append(fns, func() int { return i }) }
_ = fns }
