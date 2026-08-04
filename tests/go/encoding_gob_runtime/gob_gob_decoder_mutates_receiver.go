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
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { orig := Box{Data: []byte("ab")}
var buf bytes.Buffer
gob.NewEncoder(&buf).Encode(orig)
var back Box
gob.NewDecoder(&buf).Decode(&back)
__p(fmt.Sprint(string(back.Data))) 
__check("ab")
}
