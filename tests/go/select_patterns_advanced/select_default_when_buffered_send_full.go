// vybe-test: go/select_patterns_advanced/select_default_when_buffered_send_full
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
select { case ch <- 2: default: } }
