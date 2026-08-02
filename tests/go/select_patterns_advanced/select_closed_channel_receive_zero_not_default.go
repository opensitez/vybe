// vybe-test: go/select_patterns_advanced/select_closed_channel_receive_zero_not_default
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int)
close(ch)
select { case v := <-ch: fmt.Println(v)
default: fmt.Println("default") } }
