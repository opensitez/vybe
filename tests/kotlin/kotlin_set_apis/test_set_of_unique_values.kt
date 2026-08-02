// vybe-test: kotlin/kotlin_set_apis/test_set_of_unique_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 2, 3)
            __check((set.size).toString(), "3")
            __check((set.contains(2)).toString(), "true")
            __check((set.contains(9)).toString(), "false")
        }
