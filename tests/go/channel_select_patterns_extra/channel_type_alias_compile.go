// vybe-test: go/channel_select_patterns_extra/channel_type_alias_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
type numbers chan int
func main() { var ch numbers = make(chan int, 1)
_ = ch }
