// vybe-test: go/cover_debug_formats/buildinfo_read
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/buildinfo"
import "bytes"
func main() { _, _ = buildinfo.Read(bytes.NewReader(nil)) }
