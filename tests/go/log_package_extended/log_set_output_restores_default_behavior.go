// vybe-test: go/log_package_extended/log_set_output_restores_default_behavior
// origin: languages/go/tests/go/test_log_package_extended.rs

package main
import "fmt"
import "log"
import "bytes"
import "io"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
log.SetOutput(&buf)
log.SetFlags(0)
log.Print("captured")
log.SetOutput(io.Discard)
__check(fmt.Sprint(buf.String()), "captured\n") }
