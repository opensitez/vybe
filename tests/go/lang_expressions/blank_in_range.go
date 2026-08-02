// vybe-test: go/lang_expressions/blank_in_range
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func main() { sum := 0
for _, v := range []int{1,2} { sum += v }
fmt.Println(sum) }
