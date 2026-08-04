// vybe-test: kotlin/operators/test_vector_like_operator_overload
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Vector(val x: Int, val y: Int) {
            operator fun times(scale: Int): Vector = Vector(x * scale, y * scale)
            operator fun plus(other: Vector): Vector = Vector(x + other.x, y + other.y)
        }

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
            val a = Vector(2, 3)
            val b = a * 4
            val c = b + Vector(1, 1)
            __p((b.x).toString())
            __p((c.y).toString())
        
__check("8\n13")
}
