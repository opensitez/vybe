// vybe-test: go/channel_select_patterns_extra/channel_param_return_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func passthrough(ch chan int) chan int { return ch }
func main() { _ = passthrough }
