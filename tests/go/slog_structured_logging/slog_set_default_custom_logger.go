// vybe-test: go/slog_structured_logging/slog_set_default_custom_logger
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "bytes"
func main() { var buf bytes.Buffer
slog.SetDefault(slog.New(slog.NewTextHandler(&buf, nil)))
slog.Info("after") }
