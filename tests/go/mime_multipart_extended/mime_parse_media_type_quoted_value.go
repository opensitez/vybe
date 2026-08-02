// vybe-test: go/mime_multipart_extended/mime_parse_media_type_quoted_value
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

func main() { _, params, _ := mime.ParseMediaType(`multipart/form-data; boundary="abc123"`)
__check(fmt.Sprint(params["boundary"]), "abc123") }
