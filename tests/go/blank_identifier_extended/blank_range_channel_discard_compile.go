// vybe-test: go/blank_identifier_extended/blank_range_channel_discard_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
close(ch)
for range ch { } }
