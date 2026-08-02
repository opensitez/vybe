// vybe-test: go/encoding_gob_runtime/gob_reencode_same_type_new_buffer
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf1, buf2 bytes.Buffer
gob.NewEncoder(&buf1).Encode(88)
var v int
gob.NewDecoder(&buf1).Decode(&v)
gob.NewEncoder(&buf2).Encode(v)
var v2 int
gob.NewDecoder(&buf2).Decode(&v2)
__check(fmt.Sprint(v2), "88") }
