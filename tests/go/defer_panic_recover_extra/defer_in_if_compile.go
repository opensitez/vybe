// vybe-test: go/defer_panic_recover_extra/defer_in_if_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { if true { defer func() {}() } }
