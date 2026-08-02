// vybe-test: go/slog_structured_logging/slog_bool_value_constructor
// origin: languages/go/tests/go/test_slog_structured_logging.rs
// vybe-test-mode: compile

package main
import "log/slog"
func main() { _ = slog.Bool("k", false) }
