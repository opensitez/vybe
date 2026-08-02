// vybe-test: go/select_patterns_advanced/select_bool_channel_false_value
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan bool, 1)
ch <- false
select { case b := <-ch: fmt.Println(b)
default: fmt.Println(true) } }
