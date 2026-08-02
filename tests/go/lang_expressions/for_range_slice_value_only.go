// vybe-test: go/lang_expressions/for_range_slice_value_only
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func main() { for _, v := range []int{4} { fmt.Println(v) } }
