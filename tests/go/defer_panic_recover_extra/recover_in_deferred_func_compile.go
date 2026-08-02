// vybe-test: go/defer_panic_recover_extra/recover_in_deferred_func_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() { _ = recover() }()
panic("x") }
