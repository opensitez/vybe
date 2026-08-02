// vybe-test: go/path_filepath_package/path_join_double_dot_parent_segment
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
func main() { _ = path.Join("/a", "..", "b") }
