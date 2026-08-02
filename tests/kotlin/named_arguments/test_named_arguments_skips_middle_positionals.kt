// vybe-test: kotlin/named_arguments/test_named_arguments_skips_middle_positionals
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun score(base: Int, bonus: Int, penalty: Int): Int = base + bonus - penalty
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(base = 10, penalty = 1, bonus = 2)).toString(), "11")
        }
