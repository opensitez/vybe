// vybe-test: go/function_types_advanced/method_returns_closure_from_receiver_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type builder struct { base int }
func (b builder) incrementer() func(int) int { return func(v int) int { return b.base + v } }
func main() { _ = builder{base: 10}.incrementer() }
