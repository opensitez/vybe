// vybe-test: go/variadic_advanced/variadic_switch_on_len_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func bucket(nums ...int) int { switch len(nums) { case 0: return 0
case 1: return 1
default: return 2 } }
func main() { _ = bucket(1, 2) }
