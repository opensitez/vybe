// vybe-test: go/strings_builder/reader_read_into_buffer
// origin: languages/go/tests/go/test_strings_builder.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := strings.NewReader("abc")
buf := make([]byte, 2)
n, _ := r.Read(buf)
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(string(buf)), "ab") }
