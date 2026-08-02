// vybe-test: go/slog_structured_logging/slog_with_prepends_constant_attr
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
child := base.With("svc", "api")
child.Info("hit")
__check(fmt.Sprint(strings.Contains(buf.String(), "api")), "true") }
