// vybe-test: go/channel_direction_extended/recv_only_unbuffered_with_goroutine_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
var r <-chan int = ch
go func() { ch <- 2 }()
_ = <-r }
