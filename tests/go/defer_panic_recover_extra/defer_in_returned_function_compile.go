// vybe-test: go/defer_panic_recover_extra/defer_in_returned_function_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func build() func() { return func() { defer func() {}() } }
func main() { _ = build() }
