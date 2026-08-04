// vybe-test: go/concurrency_dummy/channel_type_send_only
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func sendData(ch chan<- int) { ch <- 1 }
func main() {}
