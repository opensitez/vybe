// vybe-test: go/channel_direction_extended/chan_direction_mismatch_param_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func use(ch chan int) {}
func main() { var s chan<- int = make(chan int, 1)
use(s) }
