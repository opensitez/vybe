// vybe-test: go/select_patterns_advanced/select_default_between_two_blocked_receives
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { left := make(chan int)
right := make(chan int)
select { case <-left: fmt.Println("left")
case <-right: fmt.Println("right")
default: fmt.Println("neither") } }
