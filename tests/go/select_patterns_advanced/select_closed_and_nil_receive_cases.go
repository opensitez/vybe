// vybe-test: go/select_patterns_advanced/select_closed_and_nil_receive_cases
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { closed := make(chan int)
close(closed)
var nilCh chan int
select { case <-closed: case <-nilCh: default: } }
