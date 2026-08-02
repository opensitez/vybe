// vybe-test: kotlin/mutable_set_apis/test_mutable_set_add_duplicate_ignored
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            val added = values.add(2)
            __check((added).toString(), "false")
            __check((values.size).toString(), "2")
        }
