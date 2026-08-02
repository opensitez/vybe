// vybe-test: go/iter_package/iter_seq_first_value_only
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
func main() { seq := func(yield func(int) bool) { yield(99)
yield(100) }
first := 0
for v := range seq { first = v
break }
fmt.Println(first) }
