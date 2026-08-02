// vybe-test: go/variadic_advanced/variadic_in_select_case_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func send(ch chan int, nums ...int) { for _, n := range nums { ch <- n } }
func main() { ch := make(chan int, 2)
send(ch, 1, 2)
_ = <-ch }
