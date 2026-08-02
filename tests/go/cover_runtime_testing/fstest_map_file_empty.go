// vybe-test: go/cover_runtime_testing/fstest_map_file_empty
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/fstest"
func main() { f := fstest.MapFile{}
_ = len(f.Data) }
