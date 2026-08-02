// vybe-test: go/iter_package/iter_seq_yield_false_halts
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
func main() { stopped := 0
seq := func(yield func(int) bool) { if !yield(1) { stopped = 1
return }
yield(2) }
for v := range seq { if v == 1 { break } }
fmt.Println(stopped) }
