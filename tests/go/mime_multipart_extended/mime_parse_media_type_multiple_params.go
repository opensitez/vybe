// vybe-test: go/mime_multipart_extended/mime_parse_media_type_multiple_params
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

func main() { mt, params, _ := mime.ParseMediaType("text/html; charset=utf-8; level=1")
__check(fmt.Sprint(mt), "text/html")
__check(fmt.Sprint(params["level"]), "1") }
