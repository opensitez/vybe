// vybe-test: go/switch_type_tagless/switch_duplicate_case_compile
// origin: languages/go/tests/go/test_switch_type_tagless.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1, 1: } }
