// vybe-test: go/mime_multipart_extended/multipart_reader_next_part
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime/multipart"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
w.CreateFormField("k")
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, err := r.NextPart()
__check(fmt.Sprint(err == nil), "true")
__check(fmt.Sprint(p != nil), "true") }
