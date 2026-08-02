// vybe-test: go/cover_text_html_log/syslog_new
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/syslog"
func main() { _, _ = syslog.New(syslog.LOG_INFO, "app") }
