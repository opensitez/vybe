// vybe-test: go/iter_package/iter_seq_range_over_custom
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
func main() { seq := func(yield func(int) bool) { for i := 1; i <= 3; i++ { if !yield(i) { return } } }
sum := 0
for v := range seq { sum += v }
fmt.Println(sum) }
