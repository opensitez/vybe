// vybe-test: go/channel_direction_extended/recv_only_with_select_ready
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 15
var r <-chan int = ch
select { case v := <-r: fmt.Println(v)
default: fmt.Println(0) } }
