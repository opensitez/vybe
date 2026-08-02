// vybe-test: go/channel_direction_extended/select_two_recv_only_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { a := make(chan int, 1)
b := make(chan int, 1)
a <- 1
b <- 2
var ra <-chan int = a
var rb <-chan int = b
select { case <-ra: case <-rb: } }
