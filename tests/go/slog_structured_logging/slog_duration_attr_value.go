// vybe-test: go/slog_structured_logging/slog_duration_attr_value
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
import "strings"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
l := slog.New(slog.NewTextHandler(&buf, nil))
l.Info("d", slog.Duration("wait", time.Second))
__check(fmt.Sprint(strings.Contains(buf.String(), "1s")), "true") }
