// vybe-test: go/channel_close_range/channel_direction_send_only
// origin: languages/go/tests/go/test_channel_close_range.rs
// vybe-test-mode: compile

package main
func send(ch chan<- int) { ch <- 1 }
func main() { ch := make(chan int, 1)
send(ch) }
