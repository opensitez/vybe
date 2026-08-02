// vybe-test: kotlin/range_apis/test_range_join_to_string
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..5
            __check((r.joinToString(".")).toString(), "1.2.3.4.5")
        }
