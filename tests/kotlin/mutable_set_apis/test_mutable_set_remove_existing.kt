// vybe-test: kotlin/mutable_set_apis/test_mutable_set_remove_existing
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            val removed = values.remove(2)
            __check((removed).toString(), "true")
            __check((values.size).toString(), "2")
        }
