// vybe-test: go/select_patterns_advanced/select_nil_plus_closed_channel_prefers_closed
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { chE := make(chan int)
close(chE)
var chN chan int
select { case <-chN: fmt.Println("nil")
case v := <-chE: fmt.Println(v)
default: fmt.Println("default") } }
