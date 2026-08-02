// vybe-test: go/defer_lifo_extended/defer_in_select_with_fallthrough_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
func main() { ch := make(chan int)
select { case <-ch: defer func() {}()
default: } }
