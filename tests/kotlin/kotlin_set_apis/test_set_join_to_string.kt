// vybe-test: kotlin/kotlin_set_apis/test_set_join_to_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = linkedSetOf("s", "t", "u")
            __check((set.joinToString("|")).toString(), "s|t|u")
        }
