// vybe-test: go/cover_image/image_new_alpha
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image"
func main() { _ = image.NewAlpha(image.Rect(0, 0, 2, 2)) }
