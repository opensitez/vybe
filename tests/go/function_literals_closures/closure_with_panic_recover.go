// vybe-test: go/function_literals_closures/closure_with_panic_recover
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func main() { safe := func() { defer func() { fmt.Println(recover() != nil) }()
panic("x") }
safe() }
