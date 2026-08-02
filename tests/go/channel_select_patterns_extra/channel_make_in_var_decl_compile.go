// vybe-test: go/channel_select_patterns_extra/channel_make_in_var_decl_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
var ch = make(chan int, 1)
func main() { _ = ch }
