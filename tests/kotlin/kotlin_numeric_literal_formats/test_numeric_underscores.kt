// vybe-test: kotlin/kotlin_numeric_literal_formats/test_numeric_underscores
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_literal_formats.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val million = 1_000_000
            val grouped = 12_34_56
            __check((million).toString(), "1000000")
            __check((grouped).toString(), "123456")
        }
