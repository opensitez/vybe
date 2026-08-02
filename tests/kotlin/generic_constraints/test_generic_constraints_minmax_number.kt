// vybe-test: kotlin/generic_constraints/test_generic_constraints_minmax_number
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> minmax(a: T, b: T): String {
            val aInt = a.toInt()
            val bInt = b.toInt()
            return if (aInt <= bInt) "$aInt:$bInt" else "$bInt:$aInt"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((minmax(7, 4)).toString(), "4:7")
            __check((minmax(2.0, 8.0)).toString(), "2:8")
        }
