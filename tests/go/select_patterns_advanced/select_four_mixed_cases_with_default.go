// vybe-test: go/select_patterns_advanced/select_four_mixed_cases_with_default
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { var nilCh chan int
ready := make(chan int, 1)
send := make(chan int, 1)
recv := make(chan int)
select { case <-nilCh: case v := <-ready: _ = v
case send <- 1: case <-recv: default: } }
