// vybe-test: go/concurrency_dummy/channel_make
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
_ = ch }
