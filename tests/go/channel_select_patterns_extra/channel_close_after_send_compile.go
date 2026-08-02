// vybe-test: go/channel_select_patterns_extra/channel_close_after_send_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
close(ch) }
