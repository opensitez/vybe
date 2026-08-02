// vybe-test: go/cover_runtime_testing/iotest_data_err_reader
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/iotest"
import "strings"
func main() { _ = iotest.DataErrReader(strings.NewReader("abc")) }
