// vybe-test: go/defer_lifo_extended/defer_register_in_init_with_panic_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func init() { defer func() { recover() }()
panic("init") }
func main() {}
