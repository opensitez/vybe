// vybe-test: go/mime_multipart_extended/mime_parse_media_type_image_png
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

func main() { mt, _, _ := mime.ParseMediaType("image/png")
__check(fmt.Sprint(mt), "image/png") }
