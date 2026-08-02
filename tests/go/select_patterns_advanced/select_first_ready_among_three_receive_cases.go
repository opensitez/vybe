// vybe-test: go/select_patterns_advanced/select_first_ready_among_three_receive_cases
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch1 := make(chan int, 1)
ch2 := make(chan int)
ch3 := make(chan int)
ch1 <- 3
select { case v := <-ch1: fmt.Println(v)
case <-ch2: fmt.Println(2)
case <-ch3: fmt.Println(1)
default: fmt.Println(0) } }
