// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_from_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf('a', 'b', 'c').associateWith { it.code }
            __check((map['a']).toString(), "97")
            __check((map['c']).toString(), "99")
        }
