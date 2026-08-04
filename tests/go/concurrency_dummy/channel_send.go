// vybe-test: go/concurrency_dummy/channel_send
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { ch <- 1 }() }
