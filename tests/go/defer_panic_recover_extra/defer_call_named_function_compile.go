// vybe-test: go/defer_panic_recover_extra/defer_call_named_function_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func cleanup() {}
func main() { defer cleanup() }
