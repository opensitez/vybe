// vybe-test: go/defer_lifo_extended/defer_after_panic_without_recover_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func run() { defer func() {}()
panic("x")
defer func() { _ = 1 }() }
func main() { run() }
