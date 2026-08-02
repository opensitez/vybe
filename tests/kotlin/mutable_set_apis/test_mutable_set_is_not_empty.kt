// vybe-test: kotlin/mutable_set_apis/test_mutable_set_is_not_empty
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf<Int>()
            __check((values.isEmpty()).toString(), "true")
            values.add(1)
            __check((values.isNotEmpty()).toString(), "true")
        }
