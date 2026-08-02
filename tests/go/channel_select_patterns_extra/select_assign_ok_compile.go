// vybe-test: go/channel_select_patterns_extra/select_assign_ok_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case v, ok := <-ch: _, _ = v, ok
default: } }
