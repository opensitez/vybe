// vybe-test: go/time_location_zone/time_until_fixed_date
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = time.Until(time.Date(2030, 1, 1, 0, 0, 0, 0, time.UTC)) }
