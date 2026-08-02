// vybe-test: go/log_package_extended/log_output_direct_call
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
log.SetFlags(0)
_ = log.Output(0, "direct\n")
log.SetOutput(&buf)
log.SetFlags(0)
log.Print("after")
__check(fmt.Sprint(buf.String()), "after\n") }
