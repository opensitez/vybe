// vybe-test: go/function_literals_closures/closure_with_defer_inside
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { run := func() { defer fmt.Println("done")
fmt.Println("go") }
run() }
