// vybe-test: go/cover_image/image_new_gray
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image"
func main() { _ = image.NewGray(image.Rect(0, 0, 2, 2)) }
