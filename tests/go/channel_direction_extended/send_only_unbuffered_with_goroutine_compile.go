// vybe-test: go/channel_direction_extended/send_only_unbuffered_with_goroutine_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
var s chan<- int = ch
go func() { s <- 1 }()
_ = <-ch }
