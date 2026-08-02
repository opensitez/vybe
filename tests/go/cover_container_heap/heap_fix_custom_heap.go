// vybe-test: go/cover_container_heap/heap_fix_custom_heap
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type kv struct { k, v int }
type kh struct { data []kv }
func (k kh) Len() int { return len(k.data) }
func (k kh) Less(i, j int) bool { return k.data[i].v < k.data[j].v }
func (k kh) Swap(i, j int) { k.data[i], k.data[j] = k.data[j], k.data[i] }
func main() { h := &kh{data: []kv{{k: 1, v: 3}, {k: 2, v: 1}}}
heap.Init(h)
h.data[0].v = 0
heap.Fix(h, 0) }
