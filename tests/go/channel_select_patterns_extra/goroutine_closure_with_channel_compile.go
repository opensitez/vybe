// vybe-test: go/channel_select_patterns_extra/goroutine_closure_with_channel_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
go func(local chan int) { local <- 1 }(ch) }
