// vybe-test: go/function_types_advanced/method_for_each_with_index_callback
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type batch struct { items []int }
func (b batch) forEach(visit func(int, int)) { for i, v := range b.items { visit(i, v) } }
func main() { sum := 0
batch{items: []int{2, 3, 4}}.forEach(func(i int, v int) { sum += v })
fmt.Println(sum) }
