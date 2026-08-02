// vybe-test: go/switch_fallthrough_extended/switch_fallthrough_in_default_only_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { default: fallthrough } }
