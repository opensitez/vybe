// vybe-test: go/select_patterns_advanced/select_string_channel_send_then_receive
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan string, 1)
select { case ch <- "go": default: }
select { case <-ch: default: } }
