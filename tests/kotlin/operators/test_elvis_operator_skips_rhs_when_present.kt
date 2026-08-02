// vybe-test: kotlin/operators/test_elvis_operator_skips_rhs_when_present
// origin: languages/kotlin/tests/kotlin/test_operators.rs

var evals = 0

        fun fallback(): Int {
            evals += 1
            return 99
        }

        fun coalesce(value: Int?): Int {
            return value ?: fallback()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((coalesce(12)).toString(), "12")
            __check((evals).toString(), "0")
            __check((coalesce(null)).toString(), "99")
            __check((evals).toString(), "1")
        }
