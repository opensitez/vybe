// vybe-test: go/channel_select_patterns_extra/select_send_only_channel_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var send chan<- int = ch
select { case send <- 1: default: } }
