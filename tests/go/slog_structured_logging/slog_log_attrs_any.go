// vybe-test: go/slog_structured_logging/slog_log_attrs_any
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { slog.LogAttrs(context.Background(), slog.LevelDebug, "any", slog.Any("v", "s")) }
