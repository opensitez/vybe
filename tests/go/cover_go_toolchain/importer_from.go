// vybe-test: go/cover_go_toolchain/importer_from
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/importer"
import "go/token"
func main() { _ = importer.From("gc", "./", nil) }
