// vybe-test: go/concurrency_dummy/select_assign
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
select { case v := <-ch: _ = v
default: } }
