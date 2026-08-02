// vybe-test: go/path_filepath_package/path_base_single_name_no_slash
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
func main() { _ = path.Base("archive.zip") }
