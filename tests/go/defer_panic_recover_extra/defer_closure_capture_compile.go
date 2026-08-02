// vybe-test: go/defer_panic_recover_extra/defer_closure_capture_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { value := 1
defer func() { _ = value }() }
