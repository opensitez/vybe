// vybe-test: go/defer_panic_recover_extra/recover_assigned_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() { value := recover()
_ = value }()
panic("x") }
