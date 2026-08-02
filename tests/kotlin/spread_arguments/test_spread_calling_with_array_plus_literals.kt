// vybe-test: kotlin/spread_arguments/test_spread_calling_with_array_plus_literals
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun show(vararg values: String): String = values.joinToString(",")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = arrayOf("b", "c")
            __check((show("a", *a, "d")).toString(), "a,b,c,d")
        }
