// vybe-test: go/mime_multipart_extended/mime_extensions_by_type_plain
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

func main() { exts, _ := mime.ExtensionsByType("text/plain")
__check(fmt.Sprint(len(exts) > 0), "true") }
