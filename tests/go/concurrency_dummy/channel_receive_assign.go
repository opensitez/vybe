// vybe-test: go/concurrency_dummy/channel_receive_assign
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
go func() { v := <-ch
_ = v }() }
