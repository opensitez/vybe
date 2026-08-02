// vybe-test: go/path_filepath_package/filepath_join_leading_absolute_segment
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Join("/tmp", "vybe", "out") }
