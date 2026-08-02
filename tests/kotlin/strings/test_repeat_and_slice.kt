// vybe-test: kotlin/strings/test_repeat_and_slice
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("ha".repeat(3)).toString(), "hahaha")
            __check(("kotlin".slice(1..3)).toString(), "otl")
            __check(("kotlin".slice(IntRange(0, 2))).toString(), "kot")
        }
