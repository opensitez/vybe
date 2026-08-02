// vybe-test: go/log_package_extended/log_prefix_and_message_order
// origin: languages/go/tests/go/test_log_package_extended.rs

package main
import "fmt"
import "log"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
log.SetOutput(&buf)
log.SetFlags(0)
log.SetPrefix(">>")
log.Print("go")
__check(fmt.Sprint(strings.HasPrefix(buf.String(), ">>")), "true") }
