// vybe-test: go/cover_runtime_testing/fstest_map_file_sys
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/fstest"
func main() { f := fstest.MapFile{}
_ = f.Sys() }
