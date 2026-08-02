// vybe-test: go/container_list_ring/ring_link_unlink
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/ring"
func main() { r := ring.New(2)
s := ring.New(2)
r.Link(s)
r.Unlink(1) }
