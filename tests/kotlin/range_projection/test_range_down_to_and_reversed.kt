// vybe-test: kotlin/range_projection/test_range_down_to_and_reversed
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 10 downTo 4
            __check((r.toList().joinToString(",")).toString(), "10,9,8,7,6,5,4")
            __check((r.first).toString(), "10")
            __check((r.last).toString(), "4")
            val asc = r.reversed()
            __check((asc.toList().joinToString(",")).toString(), "4,5,6,7,8,9,10")
        }
