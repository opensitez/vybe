// vybe-test: go/select_patterns_advanced/select_second_case_send_when_first_receive_blocked
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { recv := make(chan int)
send := make(chan int, 1)
select { case <-recv: case send <- 5: default: } }
