// vybe-test: go/slog_structured_logging/slog_log_context_background
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "context"
func main() { slog.Log(context.Background(), slog.LevelInfo, "ctx") }
