// vybe-test: go/lang_declarations_types/range_over_int_go122
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func main() { n := 0
for i := range 3 { n += i }
fmt.Println(n) }
