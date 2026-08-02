// vybe-test: go/container_heap_list/heap_empty_init
// origin: languages/go/tests/go/test_container_heap_list.rs
// vybe-test-mode: compile

package main
import "container/heap"
type IH []int
func (h IH) Len() int { return len(h) }
func (h IH) Less(i, j int) bool { return h[i] < h[j] }
func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }
func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }
func (h *IH) Pop() interface{} { o := *h
n := len(o)
x := o[n-1]
*h = o[:n-1]
return x }
func main() { h := &IH{}
heap.Init(h) }
