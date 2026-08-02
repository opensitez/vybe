// vybe-test: go/defer_lifo_extended/defer_in_defer_panic_uncaught_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func run() { defer func() { panic("inner") }()
panic("outer") }
func main() { run() }
