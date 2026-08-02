// vybe-test: go/log_package_extended/log_output_to_buffer_with_prefix
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
log.SetPrefix("P> ")
_ = log.Output(1, "msg\n")
__check(fmt.Sprint(buf.String()), "P> msg\n") }
