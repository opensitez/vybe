// vybe-test: go/channel_select_patterns_extra/channel_in_struct_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
type holder struct { ch chan int }
func main() { _ = holder{ch: make(chan int, 1)} }
