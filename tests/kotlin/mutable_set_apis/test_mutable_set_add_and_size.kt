// vybe-test: kotlin/mutable_set_apis/test_mutable_set_add_and_size
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            values.add(3)
            __check((values.size).toString(), "3")
            __check((values.contains(3)).toString(), "true")
        }
