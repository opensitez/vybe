// vybe-test: go/channel_select_patterns_extra/close_nil_channel_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var ch chan int
close(ch) }
