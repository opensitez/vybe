// vybe-test: go/channel_select_patterns_extra/select_multiple_cases_compile
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { ch1 := make(chan int, 1)
ch2 := make(chan int, 1)
select { case <-ch1: case ch2 <- 1: default: } }
