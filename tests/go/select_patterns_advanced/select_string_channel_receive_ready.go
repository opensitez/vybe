// vybe-test: go/select_patterns_advanced/select_string_channel_receive_ready
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

package main
import "fmt"
func main() { ch := make(chan string, 1)
ch <- "go"
select { case s := <-ch: fmt.Println(s)
default: fmt.Println("default") } }
