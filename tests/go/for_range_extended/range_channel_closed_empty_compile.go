// vybe-test: go/for_range_extended/range_channel_closed_empty_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan struct{})
close(ch)
for range ch { } }
