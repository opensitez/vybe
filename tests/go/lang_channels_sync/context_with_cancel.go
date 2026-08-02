// vybe-test: go/lang_channels_sync/context_with_cancel
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
import "context"
func main() { _, cancel := context.WithCancel(context.Background())
cancel() }
