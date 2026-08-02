// vybe-test: go/blank_identifier_extended/blank_range_map_key_discard_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { for _, v := range map[int]string{1: "a"} { _ = v } }
