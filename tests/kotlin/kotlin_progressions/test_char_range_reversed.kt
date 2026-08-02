// vybe-test: kotlin/kotlin_progressions/test_char_range_reversed
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = ('c' downTo 'a').toList()
            __check((out.toList().joinToString(",")).toString(), "c,b,a")
        }
