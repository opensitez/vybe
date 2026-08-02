// vybe-test: go/select_patterns_advanced/select_default_with_three_blocked_nil_cases
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { var a, b, c chan int
select { case <-a: fmt.Println(1)
case <-b: fmt.Println(2)
case <-c: fmt.Println(3)
default: fmt.Println("idle") } }
