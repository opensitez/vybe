// vybe-test: go/channel_direction_extended/cannot_close_send_only_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var s chan<- int = ch
close(s) }
