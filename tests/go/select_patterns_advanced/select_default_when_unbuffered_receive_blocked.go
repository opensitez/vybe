// vybe-test: go/select_patterns_advanced/select_default_when_unbuffered_receive_blocked
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan int)
select { case <-ch: fmt.Println("recv")
default: fmt.Println("default") } }
