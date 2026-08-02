// vybe-test: kotlin/mutable_set_apis/test_mutable_set_subtract
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mutableSetOf(1, 2, 3)
            val b = mutableSetOf(2, 4)
            val c = a - b
            __check((c.joinToString(",")).toString(), "1,3")
        }
