// vybe-test: kotlin/kotlin_set_apis/test_set_contains_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 3)
            __check((set.containsAll(listOf(1, 3))).toString(), "true")
            __check((set.containsAll(listOf(1, 4))).toString(), "false")
        }
