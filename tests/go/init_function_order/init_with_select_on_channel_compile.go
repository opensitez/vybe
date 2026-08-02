// vybe-test: go/init_function_order/init_with_select_on_channel_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
func init() { ch := make(chan int, 1)
ch <- 1
select { case v := <-ch: _ = v default: } }
func main() {}
