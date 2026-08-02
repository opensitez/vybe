// vybe-test: go/mime_multipart_extended/mime_format_media_type_two_params
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

func main() { s := mime.FormatMediaType("multipart/form-data", map[string]string{"boundary": "b", "charset": "utf-8"})
__check(fmt.Sprint(len(s) > 20), "true") }
