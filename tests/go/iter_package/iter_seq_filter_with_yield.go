// vybe-test: go/iter_package/iter_seq_filter_with_yield
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
func main() { seq := func(yield func(int) bool) { for i := 1; i <= 5; i++ { if i%2 == 0 { if !yield(i) { return } } } }
evens := 0
for range seq { evens++ }
fmt.Println(evens) }
