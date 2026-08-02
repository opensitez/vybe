// vybe-test: go/channel_select_patterns_extra/channel_in_interface_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var value interface{} = make(chan int, 1)
_ = value }
