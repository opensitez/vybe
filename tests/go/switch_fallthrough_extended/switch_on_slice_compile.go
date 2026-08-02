// vybe-test: go/switch_fallthrough_extended/switch_on_slice_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch []int{1} { case []int{1}: _ = 1 } }
