// vybe-test: go/context_package/context_passed_to_helper_function
// origin: languages/go/tests/go/test_context_package.rs
// vybe-test-mode: compile

package main
import "context"
func work(ctx context.Context) error { return ctx.Err() }
func main() { _ = work(context.Background()) }
