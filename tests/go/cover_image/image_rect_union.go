// vybe-test: go/cover_image/image_rect_union
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image"
func main() { _ = image.Rect(0, 0, 1, 1).Union(image.Rect(1, 1, 2, 2)) }
