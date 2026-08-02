// vybe-test: go/channel_direction_extended/select_send_on_recv_only_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var r <-chan int = ch
select { case r <- 1: default: } }
