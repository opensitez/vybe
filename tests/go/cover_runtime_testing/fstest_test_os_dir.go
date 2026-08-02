// vybe-test: go/cover_runtime_testing/fstest_test_os_dir
// origin: languages/go/tests/go/test_cover_runtime_testing.rs
// vybe-test-mode: compile

package main
import "testing/fstest"
import "os"
func main() { _ = fstest.TestFS(os.DirFS("."), "Cargo.toml") }
