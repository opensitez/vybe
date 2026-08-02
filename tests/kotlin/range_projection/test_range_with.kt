// vybe-test: kotlin/range_projection/test_range_with
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 1..10
            val firstHalf = r.take(3)
            val after = r.drop(7)
            __check((firstHalf.joinToString(",")).toString(), "1,2,3")
            __check((after.joinToString(",")).toString(), "8,9,10")
        }
