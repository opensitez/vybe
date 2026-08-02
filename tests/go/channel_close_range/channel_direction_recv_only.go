// vybe-test: go/channel_close_range/channel_direction_recv_only
// origin: languages/go/tests/go/test_channel_close_range.rs
// vybe-test-mode: compile

package main
func recv(ch <-chan int) int { return <-ch }
func main() { ch := make(chan int, 1)
ch <- 2
_ = recv(ch) }
