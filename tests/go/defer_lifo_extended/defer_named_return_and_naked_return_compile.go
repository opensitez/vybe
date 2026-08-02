// vybe-test: go/defer_lifo_extended/defer_named_return_and_naked_return_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func work() (n int) { defer func() { n = 1 }()
return n }
func main() { _ = work() }
