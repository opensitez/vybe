// vybe-test: go/channel_direction_extended/chan_direction_mismatch_return_compile
// origin: languages/go/tests/go/test_channel_direction_extended.rs
// vybe-test-mode: compile

package main
func bad() chan int { var s chan<- int = make(chan int, 1)
return s }
