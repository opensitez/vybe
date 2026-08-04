// vybe-test: go/mime_multipart_extended/mime_parse_media_type_simple
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

func main() { mt, params, err := mime.ParseMediaType("text/html")
__p(fmt.Sprint(mt))
__p(fmt.Sprint(err == nil))
__p(fmt.Sprint(len(params))) 
__check("text/html\ntrue\n0")
}
