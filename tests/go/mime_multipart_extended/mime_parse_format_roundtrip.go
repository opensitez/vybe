// vybe-test: go/mime_multipart_extended/mime_parse_format_roundtrip
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

func main() { orig := "text/plain; charset=utf-8"
mt, params, _ := mime.ParseMediaType(orig)
back := mime.FormatMediaType(mt, params)
__check(fmt.Sprint(back == orig), "true") }
