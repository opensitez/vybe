// vybe-test: kotlin/data_classes/test_data_class_plus_operator_style_via_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Coord(val x: Int, val y: Int)

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val origin = Coord(0, 0)
            fun move(point: Coord, dx: Int, dy: Int): Coord {
                return point.copy(x = point.x + dx, y = point.y + dy)
            }
            val moved = move(origin, 3, 4)
            __p((moved.x).toString())
            __p((moved.y).toString())
            __p((origin.x).toString())
        
__check("3\n4\n0")
}
