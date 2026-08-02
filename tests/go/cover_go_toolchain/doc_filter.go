// vybe-test: go/cover_go_toolchain/doc_filter
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/doc"
func main() { _ = doc.Filter(nil, nil) }
