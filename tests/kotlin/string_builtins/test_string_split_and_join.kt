// vybe-test: kotlin/string_builtins/test_string_split_and_join
// origin: languages/kotlin/tests/kotlin/test_string_builtins.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "a,b,c"
            __check((text.split(",").joinToString("|")).toString(), "a|b|c")
            __check((text.reversed()).toString(), "c,b,a")
        }
