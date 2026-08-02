// vybe-test: go/cover_runtime_testing/iotest_new_read_logger
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/iotest"
import "strings"
func main() { _ = iotest.NewReadLogger(strings.NewReader("x")) }
