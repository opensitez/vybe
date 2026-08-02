// vybe-test: go/channel_select_patterns_extra/channel_compare_nil_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var ch chan int
_ = (ch == nil) }
