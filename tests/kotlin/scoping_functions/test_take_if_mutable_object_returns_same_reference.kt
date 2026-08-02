// vybe-test: kotlin/scoping_functions/test_take_if_mutable_object_returns_same_reference
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box(var n: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box(3)
            val out = box.takeIf { it.n == 3 }
            __check((out === box).toString(), "true")
            __check((out?.n).toString(), "3")
        }
