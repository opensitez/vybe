// vybe-test: go/mime_multipart_extended/mime_format_media_type_empty_map
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

func main() { __check(fmt.Sprint(mime.FormatMediaType("application/octet-stream", map[string]string{})), "application/octet-stream") }
