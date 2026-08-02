// vybe-test: go/defer_panic_recover_extra/recover_without_panic_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = recover() }
