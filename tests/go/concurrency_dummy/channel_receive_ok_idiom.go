// vybe-test: go/concurrency_dummy/channel_receive_ok_idiom
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { v, ok := <-ch
_, _ = v, ok }() }
