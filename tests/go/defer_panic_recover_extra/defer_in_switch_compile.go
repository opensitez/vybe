// vybe-test: go/defer_panic_recover_extra/defer_in_switch_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1: defer func() {}() } }
