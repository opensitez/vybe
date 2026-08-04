// vybe-test: go/concurrency_dummy/channel_type_recv_only
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func recvData(ch <-chan int) { <-ch }
func main() {}
