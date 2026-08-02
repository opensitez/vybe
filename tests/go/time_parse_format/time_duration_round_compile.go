// vybe-test: go/time_parse_format/time_duration_round_compile
// origin: languages/go/tests/go/test_time_parse_format.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = (3 * time.Hour).Round(time.Minute) }
