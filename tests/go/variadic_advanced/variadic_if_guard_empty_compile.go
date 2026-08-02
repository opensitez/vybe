// vybe-test: go/variadic_advanced/variadic_if_guard_empty_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func ok(nums ...int) bool { return len(nums) > 0 }
func main() { if ok(1) { _ = ok() } }
