// vybe-test: go/mime_multipart_extended/multipart_reader_part_content_type
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime/multipart"
import "bytes"
import "net/textproto"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
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
__check(fmt.Sprint(p.Header.Get("Content-Type")), "application/json") }
