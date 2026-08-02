// vybe-test: go/cover_text_html_log/slog_with_group
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { _ = slog.WithGroup("g") }
