// vybe-test: go/select_patterns_advanced/select_default_when_nil_send_blocked
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { var ch chan int
select { case ch <- 1: fmt.Println("send")
default: fmt.Println("default") } }
