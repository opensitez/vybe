// vybe-test: go/select_patterns_advanced/select_closed_channel_receive_ok_false
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int)
close(ch)
select { case v, ok := <-ch: fmt.Println(v)
fmt.Println(ok)
default: fmt.Println("default") } }
