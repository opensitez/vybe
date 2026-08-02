// vybe-test: go/lang_functions_returns/defer_modifies_named_result
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func f() (n int) { defer func() { n++ }()
return 1 }
func main() { fmt.Println(f()) }
