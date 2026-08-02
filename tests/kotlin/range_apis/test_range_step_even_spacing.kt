// vybe-test: kotlin/range_apis/test_range_step_even_spacing
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..10).step(3)
            __check((r.first).toString(), "1")
            __check((r.last).toString(), "10")
            __check((r.step).toString(), "3")
            __check((r.toList().joinToString(",")).toString(), "1,4,7,10")
        }
