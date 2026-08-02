// vybe-test: kotlin/named_arguments/test_named_arguments_no_name_repetition
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun first(a: String, b: String): String = a + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((first(a = "a", b = "b")).toString(), "ab")
        }
