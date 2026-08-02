// vybe-test: go/channel_select_patterns_extra/channel_direction_recv_only_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func recvData(ch <-chan int) int { return <-ch }
func main() { _ = recvData }
