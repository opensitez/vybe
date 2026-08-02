// vybe-test: go/defer_panic_recover_extra/defer_inside_anonymous_function_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { func() { defer func() {}() }() }
