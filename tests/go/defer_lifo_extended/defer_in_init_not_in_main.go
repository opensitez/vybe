// vybe-test: go/defer_lifo_extended/defer_in_init_not_in_main
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
var x = func() int { defer fmt.Println("init")
return 1 }()
func main() { fmt.Println(x) }
