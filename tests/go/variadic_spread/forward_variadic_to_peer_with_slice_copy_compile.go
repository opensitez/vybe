// vybe-test: go/variadic_spread/forward_variadic_to_peer_with_slice_copy_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
func sink(nums ...int) int { return len(nums) }
func relay(nums ...int) int { copy := append([]int(nil), nums...)
return sink(copy...) }
func main() { _ = relay(3, 4, 5) }
