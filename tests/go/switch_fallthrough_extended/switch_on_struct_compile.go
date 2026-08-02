// vybe-test: go/switch_fallthrough_extended/switch_on_struct_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
type p struct { x int }
func main() { switch p{x: 1} { case p{x: 1}: _ = 1 } }
