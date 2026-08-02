// vybe-test: go/channel_select_patterns_extra/select_receive_ready_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 8
select { case v := <-ch: fmt.Println(v)
default: fmt.Println(0) } }
