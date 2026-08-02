// vybe-test: go/channel_select_patterns_extra/channel_assign_variable_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { first := make(chan int, 1)
second := first
_ = second }
