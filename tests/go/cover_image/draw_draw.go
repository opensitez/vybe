// vybe-test: go/cover_image/draw_draw
// origin: languages/go/tests/go/test_cover_image.rs
// vybe-test-mode: compile

package main
import "image/draw"
import "image"
func main() { dst := image.NewRGBA(image.Rect(0, 0, 1, 1))
draw.Draw(dst, dst.Bounds(), dst, image.Point{}, draw.Src) }
