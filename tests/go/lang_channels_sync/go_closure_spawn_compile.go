// vybe-test: go/lang_channels_sync/go_closure_spawn_compile
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
func main() { go func() {}() }
