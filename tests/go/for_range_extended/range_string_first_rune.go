// vybe-test: go/for_range_extended/range_string_first_rune
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { first := rune(0)
for _, r := range "z" { first = r
break }
fmt.Println(int(first)) }
