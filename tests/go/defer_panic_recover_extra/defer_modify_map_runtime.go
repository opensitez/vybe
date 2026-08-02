// vybe-test: go/defer_panic_recover_extra/defer_modify_map_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1}
func() { defer func() { values["a"] = 7 }() }()
fmt.Println(values["a"])
}
