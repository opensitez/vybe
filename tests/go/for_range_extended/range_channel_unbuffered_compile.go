// vybe-test: go/for_range_extended/range_channel_unbuffered_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { ch <- 1
close(ch) }()
for v := range ch { _ = v } }
