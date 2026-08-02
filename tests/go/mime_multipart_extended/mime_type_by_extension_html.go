// vybe-test: go/mime_multipart_extended/mime_type_by_extension_html
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(mime.TypeByExtension(".html")), "text/html; charset=utf-8") }
