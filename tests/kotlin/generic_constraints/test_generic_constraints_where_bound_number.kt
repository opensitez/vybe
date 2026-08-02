// vybe-test: kotlin/generic_constraints/test_generic_constraints_where_bound_number
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T> score(v: T): Int where T : Number {
            return v.toInt()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((score(7)).toString(), "7")
            __check((score(7.9)).toString(), "7")
        }
