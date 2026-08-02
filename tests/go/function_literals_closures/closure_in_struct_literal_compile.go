// vybe-test: go/function_literals_closures/closure_in_struct_literal_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
type pair struct { fn func() }
func main() { _ = pair{fn: func() {}} }
