// vybe-test: go/defer_lifo_extended/defer_prints_after_return_in_main
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { defer fmt.Println("end")
fmt.Println("start")
return
fmt.Println("skip") }
