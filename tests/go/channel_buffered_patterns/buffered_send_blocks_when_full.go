// vybe-test: go/channel_buffered_patterns/buffered_send_blocks_when_full
// origin: languages/go/tests/go/test_channel_buffered_patterns.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
ch <- 2 }
