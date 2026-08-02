// vybe-test: go/for_range_extended/range_map_int_keys_only_max
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { max := 0
for k := range map[int]string{3: "c", 7: "g", 1: "a"} { if k > max { max = k } }
fmt.Println(max) }
