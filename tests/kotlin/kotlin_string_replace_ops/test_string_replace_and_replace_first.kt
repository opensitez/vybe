// vybe-test: kotlin/kotlin_string_replace_ops/test_string_replace_and_replace_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_replace_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("aba".replace("a", "x")).toString(), "xbx")
            __check(("a1b2c".replaceFirst("1", "-")).toString(), "a-b2c")
        }
