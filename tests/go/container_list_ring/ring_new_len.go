// vybe-test: go/container_list_ring/ring_new_len
// origin: languages/go/tests/go/test_container_list_ring.rs
// vybe-test-mode: compile

package main
import "container/ring"
func main() { r := ring.New(3)
_ = r.Len() }
