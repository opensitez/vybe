// vybe-test: kotlin/generic_constraints/test_generic_constraints_number_to_byte
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : Number> asByte(v: T): Int = v.toByte().toInt()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asByte(260)).toString(), "4")
            __check((asByte(1.4)).toString(), "1")
        }
