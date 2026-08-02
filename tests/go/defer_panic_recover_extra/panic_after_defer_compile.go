// vybe-test: go/defer_panic_recover_extra/panic_after_defer_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() {}()
panic("x") }
