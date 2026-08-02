// vybe-test: kotlin/kotlin_set_apis/test_set_contains_sequence
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(5, 6, 7)
            val all = sequenceOf(5, 8, 7).all { set.contains(it) }
            __check((all).toString(), "false")
        }
