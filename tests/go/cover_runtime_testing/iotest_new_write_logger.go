// vybe-test: go/cover_runtime_testing/iotest_new_write_logger
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/iotest"
import "bytes"
func main() { _ = iotest.NewWriteLogger(bytes.NewBuffer(nil)) }
