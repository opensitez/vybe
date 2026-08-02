// vybe-test: go/iter_package/iter_seq_break_stops_yield
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
func main() { count := 0
seq := func(yield func(int) bool) { for i := 0; i < 100; i++ { if !yield(i) { return }
count++ } }
for v := range seq { if v == 2 { break } }
fmt.Println(count) }
