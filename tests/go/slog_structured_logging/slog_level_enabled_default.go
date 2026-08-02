// vybe-test: go/slog_structured_logging/slog_level_enabled_default
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { _ = slog.Default().Enabled(context.Background(), slog.LevelInfo) }
