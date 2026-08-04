// vybe-test: go/error_handling/defer_recover_pattern
// origin: languages/go/tests/go/test_error_handling.rs
// vybe-test-mode: compile

package main
import "fmt"
func safeCall() { defer func() { if r := recover(); r != nil { fmt.Println("recovered") } }()
panic("oops")
}
func main() { safeCall() }
