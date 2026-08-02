// vybe-test: kotlin/mutable_set_apis/test_mutable_set_remove_if_even_not_present
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 3, 5)
            values.removeIf { it % 2 == 0 }
            __check((values.joinToString(",")).toString(), "1,3,5")
            __check((values.size).toString(), "3")
        }
