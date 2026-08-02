// vybe-test: go/channel_direction_extended/select_default_with_directed_both_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
var s chan<- int = ch
var r <-chan int = ch
select { case s <- 1: case v := <-r: _ = v
default: } }
