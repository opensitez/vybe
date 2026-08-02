// vybe-test: kotlin/default_arguments/test_default_arguments_int_list_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun sumAll(values: List<Int> = listOf(1, 2, 3)): Int = values.sum()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumAll()).toString(), "6")
            __check((sumAll(listOf(10))).toString(), "10")
        }
