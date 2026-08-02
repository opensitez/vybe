// vybe-test: go/channel_select_patterns_extra/channel_len_builtin_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 2)
_ = len(ch) }
