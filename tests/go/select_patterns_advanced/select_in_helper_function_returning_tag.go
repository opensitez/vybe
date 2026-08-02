// vybe-test: go/select_patterns_advanced/select_in_helper_function_returning_tag
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func pick(ch chan int) int { select { case v := <-ch: return v
default: return -1 } }
func main() { ch := make(chan int, 1)
ch <- 2
_ = pick(ch) }
