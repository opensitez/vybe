// vybe-test: go/channel_select_patterns_extra/channel_direction_send_only_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func sendData(ch chan<- int) { ch <- 1 }
func main() { _ = sendData }
