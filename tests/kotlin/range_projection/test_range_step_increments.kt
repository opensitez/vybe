// vybe-test: kotlin/range_projection/test_range_step_increments
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..10 step 3)
            __check((r.toList().joinToString(",")).toString(), "1,4,7,10")
            __check((r.last).toString(), "10")
        }
