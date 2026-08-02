// vybe-test: kotlin/range_apis/test_int_range_sum_via_fold
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..4
            val total = r.fold(0) { acc, value -> acc + value }
            __check((total).toString(), "10")
        }
