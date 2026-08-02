// vybe-test: go/channel_select_patterns_extra/goroutine_receive_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { _ = <-ch }() }
