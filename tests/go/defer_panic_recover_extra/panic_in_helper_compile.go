// vybe-test: go/defer_panic_recover_extra/panic_in_helper_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func boom() { panic("x") }
func main() { defer func() { _ = recover() }()
boom() }
