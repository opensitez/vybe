// vybe-test: go/channel_close_range/range_over_closed_empty_channel
// origin: languages/go/tests/go/test_channel_close_range.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
close(ch)
for range ch { } }
