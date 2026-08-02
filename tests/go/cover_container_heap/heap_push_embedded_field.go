// vybe-test: go/cover_container_heap/heap_push_embedded_field
// origin: languages/go/tests/go/test_cover_container_heap.rs
// vybe-test-mode: compile

package main
import "container/heap"
type node struct { score int }
type wrap struct { nodes []node }
func (w wrap) Len() int { return len(w.nodes) }
func (w wrap) Less(i, j int) bool { return w.nodes[i].score < w.nodes[j].score }
func (w wrap) Swap(i, j int) { w.nodes[i], w.nodes[j] = w.nodes[j], w.nodes[i] }
func main() { h := &wrap{}
heap.Push(h, node{score: 1}) }
