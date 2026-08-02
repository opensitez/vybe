// vybe-test: go/nil_zero_semantics_extra/nil_function_guard_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var fn func()
if fn == nil { return } }
