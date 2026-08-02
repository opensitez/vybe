// vybe-test: go/control_flow_patterns_extra/select_with_receive_case_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
select { case v := <-ch: _ = v
default: } }
