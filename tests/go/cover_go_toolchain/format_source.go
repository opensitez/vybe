// vybe-test: go/cover_go_toolchain/format_source
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/format"
func main() { _, _ = format.Source([]byte("package main")) }
