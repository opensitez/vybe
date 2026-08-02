// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_prefix
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf(1, 2, 3).associateWith { "id-" + it }
            __check((map[1]).toString(), "id-1")
            __check((map[2]).toString(), "id-2")
            __check((map[3]).toString(), "id-3")
        }
