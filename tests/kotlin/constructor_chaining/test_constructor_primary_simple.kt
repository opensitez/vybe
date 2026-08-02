// vybe-test: kotlin/constructor_chaining/test_constructor_primary_simple
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Box(val v: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(1).v).toString(), "1")
        }
