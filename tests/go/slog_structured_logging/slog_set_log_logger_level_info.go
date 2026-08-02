// vybe-test: go/slog_structured_logging/slog_set_log_logger_level_info
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { _ = slog.SetLogLoggerLevel(slog.LevelInfo) }
