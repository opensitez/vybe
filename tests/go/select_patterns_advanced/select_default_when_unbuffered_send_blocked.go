// vybe-test: go/select_patterns_advanced/select_default_when_unbuffered_send_blocked
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
select { case ch <- 9: default: } }
