// vybe-test: go/embed_unsafe_size/embed_fs_var
// origin: languages/go/tests/go/test_embed_unsafe_size.rs
// vybe-test-mode: compile

package main
import "embed"
import "io/fs"
//go:embed templates/*
var tmpl embed.FS
func main() { _, _ = fs.ReadFile(tmpl, "templates/x") }
