// vybe-test: kotlin/range_apis/test_range_drop_take_like_slice
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..10
            __check((r.drop(2).take(3).joinToString(",")).toString(), "3,4,5")
        }
