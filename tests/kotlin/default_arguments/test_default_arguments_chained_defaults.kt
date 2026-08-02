// vybe-test: kotlin/default_arguments/test_default_arguments_chained_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun base(a: Int = 1): Int = a
        fun step(value: Int = base(), bonus: Int = 2): Int = value + bonus
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((step()).toString(), "3")
            __check((step(4)).toString(), "6")
            __check((step(bonus = 5)).toString(), "6")
        }
