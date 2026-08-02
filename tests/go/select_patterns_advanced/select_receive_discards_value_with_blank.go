// vybe-test: go/select_patterns_advanced/select_receive_discards_value_with_blank
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 17
select { case <-ch: fmt.Println("got")
default: fmt.Println("miss") } }
