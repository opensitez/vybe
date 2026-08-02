// vybe-test: go/encoding_gob_runtime/gob_register_name_then_decode
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Tag struct { Label string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { gob.RegisterName("TagType", Tag{})
orig := Tag{Label: "x"}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Tag
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.Label), "x") }
