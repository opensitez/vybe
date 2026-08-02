// vybe-test: go/channel_select_patterns_extra/range_over_channel_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { close(ch) }()
for value := range ch { _ = value } }
