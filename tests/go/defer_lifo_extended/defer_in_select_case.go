// vybe-test: go/defer_lifo_extended/defer_in_select_case
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 1
select { case <-ch: defer fmt.Println("sel")
default: defer fmt.Println("def") } }
