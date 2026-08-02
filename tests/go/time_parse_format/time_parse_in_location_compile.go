// vybe-test: go/time_parse_format/time_parse_in_location_compile
// origin: languages/go/tests/go/test_time_parse_format.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _, _ = time.ParseInLocation("2006-01-02", "2020-01-01", time.UTC) }
