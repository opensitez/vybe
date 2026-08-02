// vybe-test: go/lang_builtins_control/variadic_forward
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func sum(xs ...int) int { t := 0
for _, v := range xs { t += v }
return t }
func main() { fmt.Println(sum(1,2,3)) }
