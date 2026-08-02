// vybe-test: kotlin/basic/test_string_concatenation
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = "Hello"
            val second = "Kotlin"
            __check((first + ", " + second).toString(), "Hello, Kotlin")
        }
