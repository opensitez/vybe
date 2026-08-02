// vybe-test: go/range_iteration_extra/range_over_string_last_rune_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { last := rune(0)
for _, value := range "go" { last = value }
fmt.Println(int(last))
}
