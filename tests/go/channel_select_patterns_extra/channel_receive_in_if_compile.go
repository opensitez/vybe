// vybe-test: go/channel_select_patterns_extra/channel_receive_in_if_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
if value := <-ch; value >= 0 { _ = value } }
