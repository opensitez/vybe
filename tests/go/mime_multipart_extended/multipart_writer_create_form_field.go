// vybe-test: go/mime_multipart_extended/multipart_writer_create_form_field
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
p, _ := w.CreateFormField("name")
__check(fmt.Sprint(p != nil), "true")
w.Close() }
