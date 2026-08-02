// vybe-test: go/select_patterns_advanced/select_receive_ok_true_on_buffered_value
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 99
select { case v, ok := <-ch: fmt.Println(v)
fmt.Println(ok)
default: fmt.Println("default") } }
