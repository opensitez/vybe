// vybe-test: go/select_patterns_advanced/select_recv_only_typed_channel_in_case
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
var in <-chan int = ch
select { case <-in: default: } }
