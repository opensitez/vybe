// vybe-test: go/cover_text_html_log/slog_error
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { slog.Error("fail") }
