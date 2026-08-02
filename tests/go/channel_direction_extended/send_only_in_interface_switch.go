// vybe-test: go/channel_direction_extended/send_only_in_interface_switch
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case chan<- int: fmt.Println("send")
default: fmt.Println("other") } }
func main() { ch := make(chan int)
tag((chan<- int)(ch)) }
