// vybe-test: go/encoding_gob_runtime/gob_pointer_to_int_roundtrip
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

func main() { n := 17
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(&n)
var back *int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(*back), "17") }
