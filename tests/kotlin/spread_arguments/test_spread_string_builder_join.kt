// vybe-test: kotlin/spread_arguments/test_spread_string_builder_join
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parts = arrayOf("a", "b")
            val all = arrayOf("s", *parts, "t")
            __check((all.joinToString("-")).toString(), "s-a-b-t")
        }
