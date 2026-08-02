// vybe-test: go/for_range_extended/range_map_values_with_blank_key
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, v := range map[string]int{"p": 4, "q": 5} { total += v }
fmt.Println(total) }
