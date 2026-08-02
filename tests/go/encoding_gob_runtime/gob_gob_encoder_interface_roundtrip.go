// vybe-test: go/encoding_gob_runtime/gob_gob_encoder_interface_roundtrip
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Custom struct { N int }
func (c Custom) GobEncode() ([]byte, error) { return []byte{byte(c.N)}, nil }
func (c *Custom) GobDecode(b []byte) error { c.N = int(b[0])
return nil }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Custom{N: 7}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Custom
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.N), "7") }
