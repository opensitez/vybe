// vybe-test: kotlin/data_class_copying/test_data_class_component_functions_in_copy_context
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class PairValue(val a: Int, val b: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = PairValue(1, "x")
            val (a, b) = p
            val copy = p.copy(a = a + 1, b = b.uppercase())
            __check((a).toString(), "1")
            __check((b).toString(), "x")
            __check((copy.a).toString(), "2")
            __check((copy.b).toString(), "X")
        }
