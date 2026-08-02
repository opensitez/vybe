// vybe-test: kotlin/spread_arguments/test_spread_with_default_initializer_array
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun values(base: String = "v", vararg entries: Int = intArrayOf(1)): String {
            return base + entries.joinToString("|")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val preset = intArrayOf()
            __check((values()).toString(), "v1")
            __check((values("x", *preset)).toString(), "x")
        }
