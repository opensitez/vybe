// vybe-test: go/cover_hash_heap_io/text_scanner_scan
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "text/scanner"
import "strings"
func main() { var s scanner.Scanner
s.Init(strings.NewReader("a"))
_, _, _ = s.Scan() }
