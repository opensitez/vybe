// vybe-test: kotlin/generics/test_generic_where_clause_dual_constraints
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> compareAndMeasure(value: T): String
        where T : Comparable<T>, T : CharSequence {
            return value.length.toString() + ":" + if (value > value) "gt" else "eq"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((compareAndMeasure("abc")).toString(), "3:eq")
        }
