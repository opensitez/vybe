// vybe-test: kotlin/constructor_chaining/test_constructor_with_init
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Track(val x: Int) {
            val y: Int
            init { y = x * 2 }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Track(4).y).toString(), "8")
        }
