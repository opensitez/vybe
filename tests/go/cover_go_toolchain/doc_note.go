// vybe-test: go/cover_go_toolchain/doc_note
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/doc"
func main() { _ = doc.Note{Key: "BUG", Body: "x"} }
