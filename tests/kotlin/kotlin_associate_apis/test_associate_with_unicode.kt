// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_unicode
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("aa", "ß", "é").associateWith { it.length }
            __check((map["ß"]).toString(), "1")
            __check((map["é"]).toString(), "1")
        }
