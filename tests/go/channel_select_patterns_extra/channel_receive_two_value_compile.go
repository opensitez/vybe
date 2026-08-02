// vybe-test: go/channel_select_patterns_extra/channel_receive_two_value_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
value, ok := <-ch
_, _ = value, ok }
