// vybe-test: kotlin/mutable_list_apis/test_mutable_list_shuffle_is_not_stable_shape
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            val copy = values.toMutableList()
            copy.shuffle()
            __check((copy.size).toString(), "3")
            __check((copy.contains(1)).toString(), "true")
            __check((copy.contains(2)).toString(), "true")
            __check((copy.contains(3)).toString(), "true")
        }
