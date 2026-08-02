// vybe-test: go/slog_structured_logging/slog_float64_attr_value
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
l := slog.New(slog.NewTextHandler(&buf, nil))
l.Info("f", slog.Float64("x", 1.5))
__check(fmt.Sprint(strings.Contains(buf.String(), "1.5")), "true") }
