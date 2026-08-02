// vybe-test: go/encoding_gob_runtime/gob_map_string_to_struct
// origin: languages/go/tests/go/test_encoding_gob_runtime.rs

package main
import "fmt"
import "encoding/gob"
import "bytes"
type Rec struct { Val int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := map[string]Rec{"k": {Val: 9}}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back map[string]Rec
gob.NewDecoder(&buf).Decode(&back)
__check(fmt.Sprint(back["k"].Val), "9") }
