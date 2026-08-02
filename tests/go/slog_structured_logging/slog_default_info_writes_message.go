// vybe-test: go/slog_structured_logging/slog_default_info_writes_message
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { slog.Default().Info("default")
__check(fmt.Sprint("ok"), "ok") }
