// vybe-test: go/select_patterns_advanced/select_mixed_nil_and_ready_buffered_receive
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { var blocked chan int
ready := make(chan int, 1)
ready <- 6
select { case <-blocked: fmt.Println(0)
case v := <-ready: fmt.Println(v)
default: fmt.Println("default") } }
