// vybe-test: go/encoding_gob_runtime/gob_empty_string_roundtrip
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

func main() { var buf bytes.Buffer
gob.NewEncoder(&buf).Encode("")
var s string
gob.NewDecoder(&buf).Decode(&s)
__check(fmt.Sprint(len(s)), "0") }
