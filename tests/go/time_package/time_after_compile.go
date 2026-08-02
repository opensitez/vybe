// vybe-test: go/time_package/time_after_compile
// origin: languages/go/tests/go/test_time_package.rs
// vybe-test-mode: compile

package main
import "time"
func main() { _ = time.After(time.Second) }
