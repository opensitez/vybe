// vybe-test: go/slog_structured_logging/slog_log_attrs_group_value
// origin: languages/go/tests/go/test_slog_structured_logging.rs

package main
import "fmt"
import "log/slog"
import "context"
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
l.LogAttrs(context.Background(), slog.LevelInfo, "evt", slog.Group("g", slog.Int("n", 9)))
__check(fmt.Sprint(strings.Contains(buf.String(), "9")), "true") }
