// vybe-test: kotlin/default_arguments/test_default_arguments_in_middle_uses_position_and_name
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun score(base: Int, bonus: Int = 1, penalty: Int = 1): Int {
            return base + bonus - penalty
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(10)).toString(), "10")
            __check((score(10, 3)).toString(), "12")
            __check((score(10, penalty = 4)).toString(), "6")
        }
