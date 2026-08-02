// vybe-test: go/mime_multipart_extended/multipart_writer_create_part_header
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
h.Set("Content-Type", "text/plain")
p, _ := w.CreatePart(h)
__check(fmt.Sprint(p != nil), "true")
w.Close() }
