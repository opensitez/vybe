// vybe-test: kotlin/mutable_set_apis/test_mutable_set_retain_all
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3, 4)
            values.retainAll(listOf(2, 4, 6))
            __check((values.joinToString(",")).toString(), "2,4")
        }
