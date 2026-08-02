// vybe-test: go/cover_runtime_testing/iotest_truncate_writer
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/iotest"
import "bytes"
func main() { _ = iotest.TruncateWriter(bytes.NewBuffer(nil), 4) }
