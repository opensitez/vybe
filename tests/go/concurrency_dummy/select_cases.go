// vybe-test: go/concurrency_dummy/select_cases
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch1 := make(chan int)
ch2 := make(chan int)
select { case <-ch1: case ch2 <- 1: default: } }
