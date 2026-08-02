// vybe-test: go/log_package_extended/log_set_prefix_then_clear
// origin: languages/go/tests/go/test_log_package_extended.rs

package main
import "fmt"
import "log"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
log.SetOutput(&buf)
log.SetFlags(0)
log.SetPrefix("A")
log.SetPrefix("")
log.Print("z")
__check(fmt.Sprint(buf.String()), "z\n") }
