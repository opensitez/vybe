// vybe-test: go/select_patterns_advanced/select_two_ready_receives_first_case_wins
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { a := make(chan int, 1)
b := make(chan int, 1)
a <- 1
b <- 2
select { case v := <-a: fmt.Println(v)
case v := <-b: fmt.Println(v)
default: fmt.Println(0) } }
