// vybe-test: go/mime_multipart_extended/multipart_reader_next_part_form_name
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime/multipart"
import "bytes"
import "io"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
w := multipart.NewWriter(&buf)
fw, _ := w.CreateFormField("token")
fw.Write([]byte("abc"))
w.Close()
r := multipart.NewReader(&buf, w.Boundary())
p, _ := r.NextPart()
__check(fmt.Sprint(p.FormName()), "token")
b, _ := io.ReadAll(p)
__check(fmt.Sprint(string(b)), "abc") }
