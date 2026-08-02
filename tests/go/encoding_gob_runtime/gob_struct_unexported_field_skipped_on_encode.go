// vybe-test: go/encoding_gob_runtime/gob_struct_unexported_field_skipped_on_encode
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Hidden struct { Pub int
priv int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := Hidden{Pub: 5, priv: 99}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(h)
var back Hidden
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.Pub), "5")
__check(fmt.Sprint(back.priv), "0") }
