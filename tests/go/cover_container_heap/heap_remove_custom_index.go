// vybe-test: go/cover_container_heap/heap_remove_custom_index
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type slot struct { n int }
type sh struct { buf []slot }
func (s sh) Len() int { return len(s.buf) }
func (s sh) Less(i, j int) bool { return s.buf[i].n < s.buf[j].n }
func (s sh) Swap(i, j int) { s.buf[i], s.buf[j] = s.buf[j], s.buf[i] }
func main() { h := &sh{buf: []slot{{n: 3}, {n: 1}, {n: 2}}}
heap.Init(h)
_ = heap.Remove(h, 1) }
