// vybe-test: go/defer_lifo_extended/defer_named_return_struct_field
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
type pair struct { a int
b int }
func work() (p pair) { defer func() { p.b = 3 }()
p.a = 1
return }
func main() { r := work()
fmt.Println(r.a)
fmt.Println(r.b) }
