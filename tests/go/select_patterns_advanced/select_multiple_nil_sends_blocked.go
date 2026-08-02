// vybe-test: go/select_patterns_advanced/select_multiple_nil_sends_blocked
// origin: languages/go/tests/go/test_select_patterns_advanced.rs
// vybe-test-mode: compile

package main
func main() { var a, b chan int
select { case a <- 1: case b <- 2: default: } }
