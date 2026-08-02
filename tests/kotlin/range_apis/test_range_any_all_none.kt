// vybe-test: kotlin/range_apis/test_range_any_all_none
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..6
            __check((r.any { it % 2 == 0 }).toString(), "true")
            __check((r.all { it > 0 }).toString(), "true")
            __check((r.none { it < 0 }).toString(), "true")
        }
