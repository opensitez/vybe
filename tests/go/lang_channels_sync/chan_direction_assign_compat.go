// vybe-test: go/lang_channels_sync/chan_direction_assign_compat
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var send chan<- int = ch
var recv <-chan int = ch
_ = send
_ = recv }
