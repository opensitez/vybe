// vybe-test: go/select_patterns_advanced/select_receive_after_prior_send_same_channel
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case ch <- 4: default: }
select { case <-ch: default: } }
