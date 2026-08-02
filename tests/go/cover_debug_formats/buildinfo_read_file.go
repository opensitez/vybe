// vybe-test: go/cover_debug_formats/buildinfo_read_file
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/buildinfo"
func main() { _, _ = buildinfo.ReadFile("/dev/null") }
