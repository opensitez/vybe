// vybe-test: go/switch_fallthrough_extended/switch_on_map_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch map[int]int{1: 2} { case map[int]int{1: 2}: _ = 1 } }
