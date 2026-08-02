// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("cat", "dog", "eel")
            val map = words.associateWith { it.length }
            __check((map["cat"]).toString(), "3")
            __check((map["dog"]).toString(), "3")
        }
