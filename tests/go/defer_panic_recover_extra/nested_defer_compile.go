// vybe-test: go/defer_panic_recover_extra/nested_defer_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() { defer func() {}() }() }
