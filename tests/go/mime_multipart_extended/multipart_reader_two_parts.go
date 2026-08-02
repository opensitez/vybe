// vybe-test: go/mime_multipart_extended/multipart_reader_two_parts
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
w.CreateFormField("a")
w.CreateFormField("b")
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
_, _ = r.NextPart()
p2, err := r.NextPart()
__check(fmt.Sprint(err == nil), "true")
__check(fmt.Sprint(p2.FormName()), "b") }
