// vybe-test: go/channel_direction_extended/select_two_send_only_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { a := make(chan int, 1)
b := make(chan int, 1)
var sa chan<- int = a
var sb chan<- int = b
select { case sa <- 1: case sb <- 2: } }
