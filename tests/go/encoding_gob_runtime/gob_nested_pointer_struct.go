// vybe-test: go/encoding_gob_runtime/gob_nested_pointer_struct
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs
// vybe-test-mode: compile

package main
import "encoding/gob"
import "bytes"
type Node struct { Next *Node
Val int }
func main() { n := &Node{Val: 1}
_ = gob.NewEncoder(bytes.NewBuffer(nil)).Encode(n) }
