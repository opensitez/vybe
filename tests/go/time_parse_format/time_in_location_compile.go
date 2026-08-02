// vybe-test: go/time_parse_format/time_in_location_compile
// origin: languages/go/tests/go/test_time_parse_format.rs
// vybe-test-mode: compile

package main
import "time"
func main() { t := time.Now()
_ = t.In(time.UTC) }
