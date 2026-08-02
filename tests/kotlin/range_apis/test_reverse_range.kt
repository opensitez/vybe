// vybe-test: kotlin/range_apis/test_reverse_range
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..5).reversed()
            __check((r.first).toString(), "5")
            __check((r.last).toString(), "1")
            __check((r.step).toString(), "-1")
            __check((r.joinToString(",")).toString(), "5,4,3,2,1")
        }
