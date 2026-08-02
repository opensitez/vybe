// vybe-test: go/select_patterns_advanced/select_receive_wins_over_default_when_buffered_ready
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 11
select { case v := <-ch: fmt.Println(v)
default: fmt.Println(0) } }
