// vybe-test: go/select_patterns_advanced/select_two_sends_one_full_one_open
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { full := make(chan int, 1)
full <- 1
open := make(chan int, 1)
select { case full <- 2: case open <- 3: default: } }
