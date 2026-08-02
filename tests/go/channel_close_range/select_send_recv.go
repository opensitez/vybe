// vybe-test: go/channel_close_range/select_send_recv
// origin: languages/go/tests/go/test_channel_close_range.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case ch <- 1: default: } }
