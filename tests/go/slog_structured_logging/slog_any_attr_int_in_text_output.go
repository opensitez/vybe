// vybe-test: go/slog_structured_logging/slog_any_attr_int_in_text_output
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
l.Info("v", slog.Any("x", 42))
__check(fmt.Sprint(strings.Contains(buf.String(), "42")), "true") }
