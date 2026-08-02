// vybe-test: go/cover_text_html_log/syslog_priority_err
// origin: languages/go/tests/go/test_cover_text_html_log.rs
// vybe-test-mode: compile

package main
import "log/syslog"
func main() { _ = syslog.LOG_ERR }
