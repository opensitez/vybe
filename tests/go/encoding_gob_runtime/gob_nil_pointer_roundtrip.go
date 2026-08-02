// vybe-test: go/encoding_gob_runtime/gob_nil_pointer_roundtrip
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

func main() { var p *int
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(p)
var back *int
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back == nil), "true") }
