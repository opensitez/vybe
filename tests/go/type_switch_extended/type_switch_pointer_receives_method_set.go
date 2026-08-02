// vybe-test: go/type_switch_extended/type_switch_pointer_receives_method_set
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type mutator interface { Set(int) }
type cell struct { n int }
func (c *cell) Set(v int) { c.n = v }
func tag(v interface{}) { switch v.(type) { case mutator: v.(mutator).Set(5)
fmt.Println("ok")
default: fmt.Println("no") } }
func main() { c := cell{}
tag(&c) }
