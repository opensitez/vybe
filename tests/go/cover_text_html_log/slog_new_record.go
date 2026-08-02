// vybe-test: go/cover_text_html_log/slog_new_record
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
import "time"
func main() { _ = slog.NewRecord(time.Now(), slog.LevelInfo, "msg", 0) }
