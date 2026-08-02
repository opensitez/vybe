// vybe-test: go/encoding_gob_runtime/gob_gob_decoder_mutates_receiver
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Box struct { Data []byte }
func (b *Box) GobDecode(p []byte) error { b.Data = append([]byte(nil), p...)
return nil }
func (b Box) GobEncode() ([]byte, error) { return b.Data, nil }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Box{Data: []byte("ab")}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Box
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(string(back.Data)), "ab") }
