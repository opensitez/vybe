// vybe-test: kotlin/spread_arguments/test_spread_with_vararg_reference
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun total(vararg values: Int): Int {
            return values.size
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val fnRef = ::total
            val nums = intArrayOf(1,2,3)
            __check((fnRef(*nums)).toString(), "3")
        }
