// vybe-test: go/defer_panic_recover_extra/defer_in_loop_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { for i := 0; i < 2; i++ { defer func() { _ = i }() } }
