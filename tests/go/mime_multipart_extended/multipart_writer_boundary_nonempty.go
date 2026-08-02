// vybe-test: go/mime_multipart_extended/multipart_writer_boundary_nonempty
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

func main() { w := multipart.NewWriter(bytes.NewBuffer(nil))
__check(fmt.Sprint(len(w.Boundary()) > 0), "true") }
