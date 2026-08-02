// vybe-test: go/channel_buffered_patterns/buffered_recv_blocks_when_empty
// origin: languages/go/tests/go/test_channel_buffered_patterns.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
_ = <-ch }
