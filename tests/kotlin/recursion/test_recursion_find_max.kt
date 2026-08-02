// vybe-test: kotlin/recursion/test_recursion_find_max
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun maxOf(values: List<Int>): Int {
            if (values.size == 1) return values[0]
            val tail = values.drop(1)
            val candidate = maxOf(tail)
            return if (values[0] > candidate) values[0] else candidate
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((maxOf(listOf(3, 1, 9, 2))).toString(), "9")
        }
