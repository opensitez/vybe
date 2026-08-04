// vybe-test: go/concurrency_dummy/channel_range
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { for v := range ch { _ = v } }() }
