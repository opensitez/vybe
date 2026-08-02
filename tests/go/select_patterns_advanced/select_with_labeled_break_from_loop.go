// vybe-test: go/select_patterns_advanced/select_with_labeled_break_from_loop
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
loop: for { select { case <-ch: break loop
default: return } } }
