// vybe-test: go/mime_multipart_extended/multipart_writer_set_boundary
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
w.SetBoundary("customBoundary42")
__check(fmt.Sprint(w.Boundary()), "customBoundary42") }
