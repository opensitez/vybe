// vybe-test: go/lang_expressions/for_three_clause
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { if i == 1 { fmt.Println(i) } } }
