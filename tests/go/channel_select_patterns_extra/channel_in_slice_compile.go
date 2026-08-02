// vybe-test: go/channel_select_patterns_extra/channel_in_slice_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = []chan int{make(chan int, 1)} }
