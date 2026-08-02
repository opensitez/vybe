// vybe-test: go/variadic_advanced/variadic_bool_all_true
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func allTrue(flags ...bool) bool { for _, f := range flags { if !f { return false } }
return true }
func main() { fmt.Println(allTrue(true, true, true)) }
