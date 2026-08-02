// vybe-test: go/channel_select_patterns_extra/channel_select_named_case_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case value := <-ch: _ = value
default: } }
