// vybe-test: go/variadic_advanced/variadic_append_to_nil_from_variadic_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func gather(nums ...int) []int { return append([]int(nil), nums...) }
func main() { _ = gather(3, 4)[0] }
