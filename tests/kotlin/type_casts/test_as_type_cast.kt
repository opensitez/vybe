// vybe-test: kotlin/type_casts/test_as_type_cast
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val obj = "kotlin language"
            val text = obj as String
            __check((text).toString(), "kotlin language")
        }
