// vybe-test: go/cover_container_heap/heap_pop_from_named_heap
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type pair struct { a, b int }
type hp struct { vals []pair }
func (h hp) Len() int { return len(h.vals) }
func (h hp) Less(i, j int) bool { return h.vals[i].a < h.vals[j].a }
func (h hp) Swap(i, j int) { h.vals[i], h.vals[j] = h.vals[j], h.vals[i] }
func main() { h := &hp{vals: []pair{{a: 2}, {a: 1}}}
heap.Init(h)
_ = heap.Pop(h) }
