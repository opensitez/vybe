// vybe-test: kotlin/local_classes/test_local_enum_like_simple
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            enum class Mode { A, B, C }
            __check((Mode.B.name).toString(), "B")
        }
