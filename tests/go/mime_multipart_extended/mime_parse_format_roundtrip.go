// vybe-test: go/mime_multipart_extended/mime_parse_format_roundtrip
// origin: languages/go/tests/go/test_mime_multipart_extended.rs

package main
import "fmt"
import "mime"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { orig := "text/plain; charset=utf-8"
mt, params, _ := mime.ParseMediaType(orig)
back := mime.FormatMediaType(mt, params)
__p(fmt.Sprint(back == orig)) 
__check("true")
}
