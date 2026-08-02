// vybe-test: go/encoding_gob_runtime/gob_slice_of_structs
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Node struct { ID int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := []Node{{ID: 1}, {ID: 2}}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back []Node
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back[1].ID), "2") }
