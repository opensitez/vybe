// vybe-test: go/init_function_order/init_channel_make_buffered_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
var ch chan int
func init() { ch = make(chan int, 2)
ch <- 1 }
func main() { _ = <-ch }
