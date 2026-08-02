// vybe-test: go/time_location_zone/time_date_all_months
// origin: languages/go/tests/go/test_time_location_zone.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = time.Date(2020, time.February, 1, 0, 0, 0, 0, time.UTC)
_ = time.Date(2020, time.September, 1, 0, 0, 0, 0, time.UTC) }
