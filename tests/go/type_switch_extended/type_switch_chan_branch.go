// vybe-test: go/type_switch_extended/type_switch_chan_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case chan int: fmt.Println("chan") default: fmt.Println("other") } }
func main() { tag(make(chan int)) }
