// vybe-test: go/container_heap_list/ring_do_sums_values
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/ring"
func main() { r := ring.New(4)
sum := 0
for i := 0; i < 4; i++ { r.Value = i + 1
r = r.Next() }
r.Do(func(v interface{}) { sum += v.(int) })
fmt.Println(sum) }
