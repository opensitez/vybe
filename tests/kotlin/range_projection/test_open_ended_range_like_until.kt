// vybe-test: kotlin/range_projection/test_open_ended_range_like_until
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1 until 4
            __check((r.toList().joinToString(",")).toString(), "1,2,3")
            val step = (1 until 9 step 2)
            __check((step.joinToString(",")).toString(), "1,3,5,7")
        }
