// vybe-test: go/slog_structured_logging/slog_with_group_prefixes_key
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
base := slog.New(slog.NewTextHandler(&buf, nil))
child := base.WithGroup("req")
child.Info("in", slog.Int("id", 1))
__check(fmt.Sprint(strings.Contains(buf.String(), "req")), "true") }
