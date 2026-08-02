// vybe-test: go/path_filepath_package/filepath_is_abs_windows_drive_letter
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.IsAbs(`C:\Windows`) }
