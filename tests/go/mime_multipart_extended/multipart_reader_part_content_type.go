// vybe-test: go/mime_multipart_extended/multipart_reader_part_content_type
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime/multipart"
import "bytes"
import "net/textproto"
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
h := make(textproto.MIMEHeader)
h.Set("Content-Type", "application/json")
w.CreatePart(h)
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, _ := r.NextPart()
__p(fmt.Sprint(p.Header.Get("Content-Type"))) 
__check("application/json")
}
