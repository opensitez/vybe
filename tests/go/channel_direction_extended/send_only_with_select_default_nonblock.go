// vybe-test: go/channel_direction_extended/send_only_with_select_default_nonblock
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
var s chan<- int = ch
select { case s <- 1: fmt.Println("sent")
default: fmt.Println("def") } }
