// vybe-test: go/channel_select_patterns_extra/select_send_ready_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
select { case ch <- 5: fmt.Println(len(ch))
default: fmt.Println(0) } }
