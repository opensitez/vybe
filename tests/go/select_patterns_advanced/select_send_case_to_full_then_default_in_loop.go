// vybe-test: go/select_patterns_advanced/select_send_case_to_full_then_default_in_loop
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int, 1)
ch <- 1
for i := 0; i < 3; i++ { select { case ch <- i: default: } } }
