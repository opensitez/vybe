// vybe-test: go/encoding_gob_runtime/gob_struct_with_bool_field
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Flags struct { Ok bool
Count int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := Flags{Ok: true, Count: 2}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Flags
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back.Ok), "true")
__check(fmt.Sprint(back.Count), "2") }
