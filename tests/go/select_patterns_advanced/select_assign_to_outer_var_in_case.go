// vybe-test: go/select_patterns_advanced/select_assign_to_outer_var_in_case
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 5
var result int
select { case v := <-ch: result = v
default: }
_ = result }
