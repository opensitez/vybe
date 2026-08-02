// vybe-test: go/channel_select_patterns_extra/select_with_break_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case <-ch: break
default: } }
