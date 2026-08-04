// vybe-test: go/mime_multipart_extended/multipart_reader_next_part
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime/multipart"
import "bytes"
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

func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.CreateFormField("k")
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, err := r.NextPart()
__p(fmt.Sprint(err == nil))
__p(fmt.Sprint(p != nil)) 
__check("true\ntrue")
}
