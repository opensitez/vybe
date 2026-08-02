// vybe-test: kotlin/constructor_chaining/test_constructor_array_wrapping
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Arr(val items: IntArray)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Arr(intArrayOf(1, 2, 3))
            __check((a.items[1]).toString(), "2")
        }
