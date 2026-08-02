// vybe-test: go/time_location_zone/time_truncate_nanosecond
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { t := time.Now()
_ = t.Truncate(time.Nanosecond) }
