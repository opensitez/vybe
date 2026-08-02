// vybe-test: go/slog_structured_logging/slog_new_record_time
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "time"
func main() { _ = slog.NewRecord(time.Now(), slog.LevelInfo, "rec", 0) }
