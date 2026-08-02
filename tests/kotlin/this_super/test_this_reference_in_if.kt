// vybe-test: kotlin/this_super/test_this_reference_in_if
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

class Holder {
        fun value(v: Int?): Int {
            return if (this.hashCode() > 0) (v ?: 0) + 1 else 0
        }
    }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Holder().value(3)).toString(), "4") }
