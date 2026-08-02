// vybe-test: kotlin/operators/test_vector_like_operator_overload
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Vector(val x: Int, val y: Int) {
            operator fun times(scale: Int): Vector = Vector(x * scale, y * scale)
            operator fun plus(other: Vector): Vector = Vector(x + other.x, y + other.y)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Vector(2, 3)
            val b = a * 4
            val c = b + Vector(1, 1)
            __check((b.x).toString(), "8")
            __check((c.y).toString(), "13")
        }
