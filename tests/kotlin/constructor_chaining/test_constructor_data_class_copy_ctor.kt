// vybe-test: kotlin/constructor_chaining/test_constructor_data_class_copy_ctor
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

data class P(val a: Int, val b: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = P(1, "x")
            val c = p.copy(a = 2)
            __check((c.a).toString(), "2")
            __check((c.b).toString(), "x")
        }
