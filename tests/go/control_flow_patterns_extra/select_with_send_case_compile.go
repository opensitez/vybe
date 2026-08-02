// vybe-test: go/control_flow_patterns_extra/select_with_send_case_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
select { case ch <- 1: default: } }
