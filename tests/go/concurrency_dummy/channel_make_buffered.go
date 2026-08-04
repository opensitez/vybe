// vybe-test: go/concurrency_dummy/channel_make_buffered
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 10)
_ = ch }
